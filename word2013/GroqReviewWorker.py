# Haronex Groq Qwen Direct API worker. No browser, no server, no key in EXE.
# Protocol: stdin JSON including the key, stdout single JSON. Do not log key/report.
import json
import sys
import urllib.request
import urllib.error
import time
import email.utils
import datetime

MODEL = "qwen/qwen3.8-27b"
RETRY_DELAYS = (2, 5, 10)  # At most 3 retries; never retry generic HTTP 400.
MAX_COMPLETION_TOKENS = 8192
MIN_SPLIT_CHARS = 600
MAX_SPLIT_DEPTH = 3
FORMAT_RETRIES = 1
DISCARDED_ISSUES = {"count": 0}

ENDPOINT = "https://api.groq.com/openai/v1/chat/completions"
SYSTEM = (
    "You are Haronex, a careful English radiology report proofreader. "
    "Prioritize clinically meaningful INTERNAL contradictions before language errors. "
    "Compare Findings with Conclusion, laterality, negations, counts, anatomy, dates, "
    "measurements and clinical meaning. Distinguish metabolic versus anatomical "
    "response; residual non-avid lesions do not by themselves contradict complete "
    "metabolic response. Incidental findings omitted from an impression are not "
    "automatically contradictions; if a conclusion claims the entire remaining "
    "examination is unremarkable, flag a clearly documented remaining abnormality "
    "as a possible omission requiring clinical judgment. "
    "Do not treat local abbreviations such as IHBR as definitely incorrect. "
    "Then proofread spelling, grammar, punctuation and awkward constructions. "
    "Only flag well-supported errors. Never invent imaging findings or diagnoses. "
    "Quotes MUST be exact substrings of supplied report and no longer than 500 "
    "characters. Use category logic only for potentially clinically significant "
    "internal contradictions, and language for wording concerns. "
    "Do not repeat patient identifiers in the explanation or suggestion. "
    "Return only JSON: {\"issues\":[{\"quote\":\"exact report text\","
    "\"category\":\"logic or language\",\"explanation\":\"brief explanation\","
    "\"suggestion\":\"optional proposed wording\"}]}. "
    "Keep suggestions clinically conservative; if uncertain, recommend verification."
)

def normalize(text, issues):
    output, seen = [], set()
    if not isinstance(issues, list):
        raise ValueError("Groq did not return issues array")
    for item in issues:
        if not isinstance(item, dict):
            raise ValueError("Invalid issue item")
        quote = item.get("quote")
        category = item.get("category")
        if not isinstance(quote, str) or not 2 <= len(quote) <= 500:
            raise ValueError("Invalid quote")
        if quote not in text or category not in ("logic", "language"):
            raise ValueError("Invalid quote location or category")
        explanation = item.get("explanation", "")
        suggestion = item.get("suggestion", "")
        if not isinstance(explanation, str) or not isinstance(suggestion, str):
            raise ValueError("Invalid issue wording")
        identity = (quote, category)
        if identity in seen:
            continue
        seen.add(identity)
        output.append({"quote": quote, "category": category,
                       "explanation": explanation[:500], "suggestion": suggestion[:250]})
    return sorted(output, key=lambda x: (text.find(x["quote"]), x["category"]))

def normalize_response(text, issues):
    """Tolerate bad individual rows, but never turn an entirely invalid review into success.

    Quotes must be exact substrings. No autocorrection, case folding, invented
    text, or modifications to clinical statements. Only fixed diagnostic codes.
    """
    if not isinstance(issues, list):
        raise ReviewFailure("GROQ_ISSUES_NOT_ARRAY")
    clean, seen = [], set()
    invalid = []
    for item in issues:
        if not isinstance(item, dict):
            invalid.append("GROQ_INVALID_ITEM")
            continue
        quote, category = item.get("quote"), item.get("category")
        if not isinstance(quote, str) or not 2 <= len(quote) <= 500 or quote not in text:
            invalid.append("GROQ_INVALID_QUOTE")
            continue
        if category not in ("logic", "language"):
            invalid.append("GROQ_INVALID_CATEGORY")
            continue
        explanation, suggestion = item.get("explanation", ""), item.get("suggestion", "")
        if not isinstance(explanation, str) or not isinstance(suggestion, str):
            invalid.append("GROQ_INVALID_WORDING")
            continue
        key = (quote, category)
        if key in seen:
            continue
        seen.add(key)
        clean.append({"quote": quote, "category": category,
                      "explanation": explanation[:500], "suggestion": suggestion[:250]})
    # Empty issues [] is a valid review response. Nonempty but entirely invalid is not.
    if issues and not clean:
        raise ReviewFailure(invalid[0] if invalid else "GROQ_ALL_ISSUES_INVALID")
    DISCARDED_ISSUES["count"] += len(invalid)
    return sorted(clean, key=lambda x: (text.find(x["quote"]), x["category"]))


def parse_review_response(raw, text):
    """Classify API envelope, truncation, JSON and schema without logging raw data."""
    try:
        obj = json.loads(raw.decode("utf-8"))
    except (ValueError, UnicodeError, AttributeError):
        raise ReviewFailure("GROQ_INVALID_ENVELOPE")
    if not isinstance(obj, dict) or not isinstance(obj.get("choices"), list) or not obj["choices"]:
        raise ReviewFailure("GROQ_NO_CHOICES")
    choice = obj["choices"][0]
    if not isinstance(choice, dict):
        raise ReviewFailure("GROQ_INVALID_ENVELOPE")
    finish = choice.get("finish_reason")
    if finish == "length":
        raise ReviewFailure("GROQ_TRUNCATED_LENGTH")
    if finish == "content_filter":
        raise ReviewFailure("GROQ_CONTENT_FILTER")
    if finish != "stop":
        raise ReviewFailure("GROQ_FINISH_OTHER")
    message = choice.get("message")
    if not isinstance(message, dict):
        raise ReviewFailure("GROQ_INVALID_ENVELOPE")
    content = message.get("content")
    if not isinstance(content, str) or not content.strip():
        raise ReviewFailure("GROQ_EMPTY_RESPONSE")
    try:
        parsed = json.loads(content)
    except ValueError:
        raise ReviewFailure("GROQ_INVALID_JSON")
    if not isinstance(parsed, dict):
        raise ReviewFailure("GROQ_INVALID_STRUCTURE")
    return normalize_response(text, parsed.get("issues"))


def classify_http_400(exc):
    # Groq may return descriptions containing a user's report excerpt.
    # Inspect in memory, return only a fixed code. Never print/store raw HTTP bodies.
    try:
        body = exc.read(16384).decode("utf-8", "replace")
        parsed = json.loads(body)
        error = parsed.get("error", {})
        if not isinstance(error, dict):
            error = {}
        diagnostic = (str(error.get("code", "")) + " " +
                      str(error.get("type", "")) + " " +
                      str(error.get("message", ""))).lower()
    except Exception:
        return "GROQ_REQUEST_400_UNCLASSIFIED"
    if "response_format" in diagnostic or "json_object" in diagnostic or "json_schema" in diagnostic:
        return "GROQ_REQUEST_400_JSON_FORMAT"
    if "reasoning" in diagnostic:
        return "GROQ_REQUEST_400_REASONING"
    if "temperature" in diagnostic or "parameter" in diagnostic or "unsupported" in diagnostic:
        return "GROQ_REQUEST_400_PARAMETER"
    if "context_length" in diagnostic or "context window" in diagnostic or "too many tokens" in diagnostic:
        return "GROQ_REQUEST_400_CONTEXT"
    if "max_completion_tokens" in diagnostic or "max_tokens" in diagnostic:
        return "GROQ_REQUEST_400_MAX_TOKENS"
    if "content_filter" in diagnostic or "safety" in diagnostic or "policy" in diagnostic:
        return "GROQ_REQUEST_400_CONTENT"
    if "invalid_request_error" in diagnostic:
        return "GROQ_REQUEST_400_INVALID"
    return "GROQ_REQUEST_400_OTHER"


class ReviewFailure(Exception):
    """Safe machine-readable category. No report, key, headers or server body."""
    def __init__(self, code, retryable=False, retry_after=None):
        super().__init__(code)
        self.code = code
        self.retryable = retryable
        self.retry_after = retry_after


def retry_after_seconds(headers):
    """Honor short rate-limit Retry-After without waiting unboundedly."""
    try:
        value = headers.get("Retry-After")
        if not value:
            return None
        try:
            return max(0, min(30, float(value)))
        except (ValueError, TypeError):
            when = email.utils.parsedate_to_datetime(value)
            now = datetime.datetime.now(datetime.timezone.utc)
            return max(0, min(30, (when - now).total_seconds()))
    except Exception:
        return None


def transport_request(req, open_url=urllib.request.urlopen):
    try:
        with open_url(req, timeout=45) as response:
            return response.read(3 * 1024 * 1024)
    except urllib.error.HTTPError as exc:
        code = exc.code
        if code == 400:
            raise ReviewFailure(classify_http_400(exc))
        if code == 401:
            raise ReviewFailure("GROQ_AUTH_401")
        if code == 403:
            raise ReviewFailure("GROQ_FORBIDDEN_403")
        if code == 404:
            raise ReviewFailure("GROQ_MODEL_404")
        if code == 408 or code == 429:
            raise ReviewFailure("GROQ_HTTP_" + str(code), True, retry_after_seconds(exc.headers))
        if code in (500, 502, 503, 504):
            raise ReviewFailure("GROQ_HTTP_" + str(code), True, retry_after_seconds(exc.headers))
        raise ReviewFailure("GROQ_HTTP_" + str(code))
    except (urllib.error.URLError, TimeoutError, ConnectionError):
        raise ReviewFailure("GROQ_NETWORK", True)


def api_review(key, text, requester=transport_request, sleeper=time.sleep):
    request_body = {
        "model": MODEL,
        "messages": [{"role": "system", "content": SYSTEM},
                     {"role": "user", "content": "Review this report completely:\\n\\n" + text}],
        "response_format": {"type": "json_object"},
        "reasoning_effort": "low",
        "reasoning_format": "hidden",
        "temperature": 0.2,
        "max_completion_tokens": MAX_COMPLETION_TOKENS
    }
    data = json.dumps(request_body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(ENDPOINT, data=data, method="POST",
        headers={"Authorization": "Bearer " + key,
                 "Content-Type": "application/json",
                 "Accept": "application/json",
                 "User-Agent": "Haronex-Radiology-Agent/9.3"})
    # Semantic format errors get only one additional fresh response.
    # HTTP 400/auth/content filtering or clinical safety failures are never retried.
    recoverable_formats = {
        "GROQ_INVALID_ENVELOPE", "GROQ_INVALID_JSON",
        "GROQ_INVALID_STRUCTURE", "GROQ_ISSUES_NOT_ARRAY",
        "GROQ_INVALID_QUOTE", "GROQ_INVALID_CATEGORY",
        "GROQ_INVALID_WORDING", "GROQ_INVALID_ITEM",
        "GROQ_ALL_ISSUES_INVALID", "GROQ_EMPTY_RESPONSE",
        "GROQ_NO_CHOICES"
    }
    for format_attempt in range(FORMAT_RETRIES + 1):
        for attempt in range(len(RETRY_DELAYS) + 1):
            try:
                raw = requester(req)
                break
            except ReviewFailure as exc:
                if not exc.retryable or attempt >= len(RETRY_DELAYS):
                    raise
                delay = RETRY_DELAYS[attempt]
                if exc.retry_after is not None:
                    delay = max(delay, exc.retry_after)
                sleeper(delay)
        try:
            return parse_review_response(raw, text)
        except ReviewFailure as exc:
            if exc.code not in recoverable_formats or format_attempt >= FORMAT_RETRIES:
                raise
            sleeper(1)
    raise ReviewFailure("GROQ_INVALID_RESPONSE")

def split_report_text(text):
    """Lossless, bounded split, prefer paragraph boundary, never strip characters."""
    if len(text) < MIN_SPLIT_CHARS * 2:
        return None
    mid = len(text) // 2
    search_low = max(MIN_SPLIT_CHARS, mid - len(text) // 6)
    search_high = min(len(text) - MIN_SPLIT_CHARS, mid + len(text) // 6)
    candidates = [x + 1 for x in range(search_low, search_high)
                  if text[x] in "\r\n"]
    boundary = min(candidates, key=lambda x: abs(x - mid)) if candidates else mid
    if boundary < MIN_SPLIT_CHARS or len(text) - boundary < MIN_SPLIT_CHARS:
        return None
    left, right = text[:boundary], text[boundary:]
    assert left + right == text
    return left, right


def review_with_recovery(key, text, review=api_review, depth=0):
    """On finish_reason=length, split only the affected chunk; never use partial JSON.
    Aggregation is transactional: exception from any child discards all results.
    """
    try:
        return review(key, text)
    except ReviewFailure as exc:
        if exc.code != "GROQ_TRUNCATED_LENGTH" or depth >= MAX_SPLIT_DEPTH:
            raise
        pieces = split_report_text(text)
        if pieces is None:
            raise
        all_issues = []
        for part in pieces:
            all_issues.extend(review_with_recovery(key, part, review, depth + 1))
        # Validate against original text and deduplicate repeated anchored quotes.
        return normalize(text, all_issues)

def main():
    try:
        incoming = json.loads(sys.stdin.read())
        key = incoming.get("api_key")
        texts = incoming.get("texts")
        version = incoming.get("review_version")
        if not isinstance(key, str) or not key.startswith("gsk_") or len(key) < 20 or len(key) > 256:
            raise ValueError("Missing or invalid Groq API key")
        if not isinstance(texts, list) or not 1 <= len(texts) <= 40:
            raise ValueError("Invalid report parts")
        if any(not isinstance(x, str) or not x or len(x) > 6000 for x in texts):
            raise ValueError("Invalid report text")
        if version != "haronex-groq-qwen38-v9-3":
            raise ValueError("Unsupported review version")
        groups = [review_with_recovery(key, text) for text in texts]
        print(json.dumps({"ok": True, "mode": "groq", "passes": 1,
                          "review_version": version, "groups": groups,
                          "discarded_issues": DISCARDED_ISSUES["count"]},
                         ensure_ascii=False))
    except Exception as exc:
        # No error bodies, secrets or reports in stdout.
        # Emit only fixed error categories; never expose raw server response or report.
        message = str(exc)
        safe_codes = {"GROQ_AUTH_401", "GROQ_FORBIDDEN_403", "GROQ_MODEL_404",
                      "GROQ_NETWORK", "GROQ_TRUNCATED_LENGTH", "GROQ_FINISH_OTHER",
                      "GROQ_CONTENT_FILTER", "GROQ_EMPTY_RESPONSE",
                      "GROQ_INVALID_JSON", "GROQ_INVALID_STRUCTURE",
                      "GROQ_INVALID_RESPONSE", "GROQ_NO_CHOICES", "GROQ_INVALID_ENVELOPE",
                      "GROQ_ISSUES_NOT_ARRAY", "GROQ_INVALID_ITEM",
                      "GROQ_INVALID_QUOTE", "GROQ_INVALID_CATEGORY",
                      "GROQ_INVALID_WORDING", "GROQ_ALL_ISSUES_INVALID",
                      "GROQ_HTTP_408", "GROQ_HTTP_429", "GROQ_HTTP_500",
                      "GROQ_HTTP_502", "GROQ_HTTP_503", "GROQ_HTTP_504"}
        code = message if (message in safe_codes or message.startswith("GROQ_REQUEST_400_")) else "WORKER_" + type(exc).__name__.upper()
        print(json.dumps({"ok": False, "error": code}))

def self_test():
    sample = "Mild fatty hepatomegaly. Small fat containing umbilical hernia."
    issues = normalize(sample, [{"quote": "Small fat containing umbilical hernia",
                                "category": "language", "explanation": "Missing punctuation",
                                "suggestion": "Small fat-containing umbilical hernia."}])
    assert len(issues) == 1 and issues[0]["category"] == "language"
    assert MODEL == "qwen/qwen3.8-27b"
    class FakeError:
        def read(self, limit):
            return b'{"error":{"message":"unsupported reasoning effort"}}'
    assert classify_http_400(FakeError()) == "GROQ_REQUEST_400_REASONING"
    assert "Haronex-Radiology-Agent/9.3" in str(api_review.__code__.co_consts)
    assert MAX_COMPLETION_TOKENS == 8192
    try:
        normalize(sample, [{"quote": "not present", "category": "logic"}])
    except ValueError:
        pass
    else:
        raise AssertionError("Invalid exact quotation was not rejected")
    # Offline regression: success after two transient failures and no 400 retry.
    import io
    from email.message import Message
    fake_response = json.dumps({"choices": [{"finish_reason": "stop", "message": {
        "content": json.dumps({"issues": []})}}]}).encode("utf-8")
    called, delays = [], []
    def flaky(request):
        called.append(1)
        if len(called) < 3:
            raise ReviewFailure("GROQ_HTTP_503", True)
        return fake_response
    assert api_review("gsk_dummy", sample, flaky, delays.append) == []
    assert len(called) == 3 and delays == [2, 5]
    called.clear()
    def invalid(request):
        called.append(1)
        raise ReviewFailure("GROQ_REQUEST_400_INVALID")
    try:
        api_review("gsk_dummy", sample, invalid, delays.append)
        raise AssertionError("400 error was silently swallowed")
    except ReviewFailure as exc:
        assert exc.code == "GROQ_REQUEST_400_INVALID" and len(called) == 1
    called.clear()
    def limited(request):
        called.append(1)
        raise ReviewFailure("GROQ_HTTP_429", True, 4)
    delays.clear()
    try:
        api_review("gsk_dummy", sample, limited, delays.append)
        raise AssertionError("retry limit was not enforced")
    except ReviewFailure:
        assert len(called) == 4 and delays == [4, 5, 10]
    assert retry_after_seconds({"Retry-After": "999"}) == 30
    # Adaptive split: failed parent, two successful halves. Preserve exact text.
    long_sample = ("FINDINGS: Normal right kidney.\n" * 70) + "IMPRESSION: Normal right kidney."
    seen = []
    def truncate_then_recover(key, segment):
        seen.append(segment)
        if len(segment) > 1200:
            raise ReviewFailure("GROQ_TRUNCATED_LENGTH")
        return []
    assert review_with_recovery("gsk_dummy", long_sample, truncate_then_recover) == []
    assert seen[0] == long_sample and len(seen) > 1
    assert split_report_text(long_sample)[0] + split_report_text(long_sample)[1] == long_sample
    def always_truncate(key, segment):
        raise ReviewFailure("GROQ_TRUNCATED_LENGTH")
    try:
        review_with_recovery("gsk_dummy", long_sample, always_truncate)
        raise AssertionError("Truncation must not be marked successful")
    except ReviewFailure as exc:
        assert exc.code == "GROQ_TRUNCATED_LENGTH"
    def fail_second_half(key, segment):
        if len(segment) > 1200:
            raise ReviewFailure("GROQ_TRUNCATED_LENGTH")
        if "IMPRESSION:" in segment:
            raise ReviewFailure("GROQ_AUTH_401")
        return []
    try:
        review_with_recovery("gsk_dummy", long_sample, fail_second_half)
        raise AssertionError("Partial results must not count as success")
    except ReviewFailure as exc:
        assert exc.code == "GROQ_AUTH_401"
    assert split_report_text("Tiny sample") is None

    # A valid issue survives a malformed neighbor; no fabricated corrections.
    good = {"quote": "Mild fatty hepatomegaly", "category": "language",
            "explanation": "Wording", "suggestion": "Consider style review"}
    bad_quote = {"quote": "invented lesion", "category": "logic",
                 "explanation": "Not in report", "suggestion": ""}
    assert normalize_response(sample, [good, bad_quote]) == [good]
    assert DISCARDED_ISSUES["count"] >= 1
    assert normalize_response(sample, []) == []
    for rejected, code in [
        ([bad_quote], "GROQ_INVALID_QUOTE"),
        ([{"quote": "Mild fatty hepatomegaly", "category": "unknown"}], "GROQ_INVALID_CATEGORY"),
        ([{"quote": "Mild fatty hepatomegaly", "category": "language",
           "explanation": []}], "GROQ_INVALID_WORDING"),
        ([False], "GROQ_INVALID_ITEM")]:
        try:
            normalize_response(sample, rejected)
            raise AssertionError("Invalid-only response marked successful")
        except ReviewFailure as exc:
            assert exc.code == code
    for invalid_content, code in [
        ('{"missing":[]}', "GROQ_ISSUES_NOT_ARRAY"),
        ('{broken', "GROQ_INVALID_JSON")]:
        packet = json.dumps({"choices":[{"finish_reason":"stop",
            "message":{"content":invalid_content}}]}).encode("utf-8")
        try:
            parse_review_response(packet, sample)
            raise AssertionError("Invalid API response marked successful")
        except ReviewFailure as exc:
            assert exc.code == code
    packet = json.dumps({"choices":[{"finish_reason":"stop",
        "message":{"content":json.dumps({"issues":[good,bad_quote]})}}]}).encode("utf-8")
    assert parse_review_response(packet, sample) == [good]
    # At most one retry for malformed JSON; no additional retry for an HTTP 400.
    calls, waits = [], []
    bad_packet = json.dumps({"choices":[{"finish_reason":"stop",
        "message":{"content":"not-json"}}]}).encode("utf-8")
    def repaired(req):
        calls.append(1)
        return bad_packet if len(calls) == 1 else packet
    assert api_review("gsk_dummy", sample, repaired, waits.append) == [good]
    assert len(calls) == 2 and waits == [1]
    calls.clear()
    def malformed(req):
        calls.append(1)
        return bad_packet
    try:
        api_review("gsk_dummy", sample, malformed, lambda _: None)
        raise AssertionError("Repeated malformed API reply marked successful")
    except ReviewFailure as exc:
        assert exc.code == "GROQ_INVALID_JSON" and len(calls) == 2
    print(json.dumps({"ok": True, "test": "groq_parser_offline"}))

if __name__ == "__main__":
    if "--self-test" in sys.argv:
        self_test()
    else:
        main()
