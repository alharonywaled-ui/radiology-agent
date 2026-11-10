# Haronex Groq Qwen Direct API worker. No browser, no server, no key in EXE.
# Protocol: stdin JSON including the key, stdout single JSON. Do not log key/report.
import json
import sys
import urllib.request
import urllib.error

MODEL = "qwen/qwen3.8-27b"
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


def api_review(key, text):
    request_body = {
        "model": MODEL,
        "messages": [{"role": "system", "content": SYSTEM},
                     {"role": "user", "content": "Review this report completely:\n\n" + text}],
        "response_format": {"type": "json_object"},
        "reasoning_effort": "low",
        "reasoning_format": "hidden",
        "temperature": 0.2,
        "max_completion_tokens": 3200
    }
    data = json.dumps(request_body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(ENDPOINT, data=data, method="POST",
        headers={"Authorization": "Bearer " + key,
                 "Content-Type": "application/json",
                 "Accept": "application/json",
                 "User-Agent": "Haronex-Radiology-Agent/7.2"})
    try:
        with urllib.request.urlopen(req, timeout=45) as response:
            raw = response.read(3 * 1024 * 1024)
    except urllib.error.HTTPError as exc:
        # Avoid logging response headers/body or access credentials.
        code = exc.code
        if code == 401:
            raise RuntimeError("GROQ_AUTH_401")
        if code == 403:
            raise RuntimeError("GROQ_FORBIDDEN_403")
        if code == 404:
            raise RuntimeError("GROQ_MODEL_404")
        if code == 429:
            raise RuntimeError("GROQ_LIMIT_429")
        if code == 400:
            raise RuntimeError(classify_http_400(exc))
        if code >= 500:
            raise RuntimeError("GROQ_SERVER_5XX")
        raise RuntimeError("GROQ_HTTP_" + str(code))
    except urllib.error.URLError:
        raise RuntimeError("GROQ_NETWORK")
    obj = json.loads(raw.decode("utf-8"))
    choices = obj.get("choices", [])
    if not choices:
        raise ValueError("Groq returned no choices")
    choice = choices[0]
    if choice.get("finish_reason") not in ("stop",):
        raise ValueError("GROQ_TRUNCATED")
    message = choice.get("message") or {}
    content = message.get("content")
    if not isinstance(content, str) or not content.strip():
        raise ValueError("GROQ_EMPTY_RESPONSE")
    try:
        parsed = json.loads(content)
    except (ValueError, TypeError):
        raise ValueError("GROQ_INVALID_JSON")
    if not isinstance(parsed, dict):
        raise ValueError("Groq returned invalid structure")
    return normalize(text, parsed.get("issues"))

def main():
    try:
        incoming = json.loads(sys.stdin.read())
        key = incoming.get("api_key")
        texts = incoming.get("texts")
        version = incoming.get("review_version")
        if not isinstance(key, str) or not key.startswith("gsk_") or len(key) < 20 or len(key) > 256:
            raise ValueError("Missing or invalid Groq API key")
        if not isinstance(texts, list) or not 1 <= len(texts) <= 20:
            raise ValueError("Invalid report parts")
        if any(not isinstance(x, str) or not x or len(x) > 12000 for x in texts):
            raise ValueError("Invalid report text")
        if version != "haronex-groq-qwen38-v7-6":
            raise ValueError("Unsupported review version")
        groups = [api_review(key, text) for text in texts]
        print(json.dumps({"ok": True, "mode": "groq", "passes": 1,
                          "review_version": version, "groups": groups},
                         ensure_ascii=False))
    except Exception as exc:
        # No error bodies, secrets or reports in stdout.
        # Emit only fixed error categories; never expose raw server response or report.
        message = str(exc)
        safe_codes = {"GROQ_AUTH_401", "GROQ_FORBIDDEN_403", "GROQ_MODEL_404",
                      "GROQ_LIMIT_429", "GROQ_REQUEST_400", "GROQ_SERVER_5XX",
                      "GROQ_NETWORK", "GROQ_TRUNCATED", "GROQ_EMPTY_RESPONSE",
                      "GROQ_INVALID_JSON"}
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
    assert "User-Agent" in api_review.__code__.co_consts or "Haronex-Radiology-Agent/7.2" in str(api_review.__code__.co_consts)
    try:
        normalize(sample, [{"quote": "not present", "category": "logic"}])
    except ValueError:
        pass
    else:
        raise AssertionError("Invalid exact quotation was not rejected")
    print(json.dumps({"ok": True, "test": "groq_parser_offline"}))

if __name__ == "__main__":
    if "--self-test" in sys.argv:
        self_test()
    else:
        main()
