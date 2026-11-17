# Haronex Groq Qwen Direct API worker. No browser, no server, no key in EXE.
# Protocol: stdin JSON including the key, stdout single JSON. Do not log key/report.
import json
import sys
import urllib.request
import urllib.error
import time
import email.utils
import datetime
import re
from concurrent.futures import ThreadPoolExecutor

MODEL = "qwen/qwen3.8-27b"
RETRY_DELAYS = (2, 6, 15)  # Bounded exponential backoff, four total HTTP attempts.
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
    "Medical abbreviations and acronyms (e.g. CT, MR, MRI, MRCP, IHBR, PET CT, BMD) "
    "are NOT automatically errors, even if uncommon. Do not flag dates or years "
    "such as 2026, 2027, or 2026/2027 merely for being recent or unexpected. "
    "Independently inspect EVERY clinical sentence, finding, and impression for "
    "contextually nonsensical or obviously unrelated inserted words, phrases, "
    "and statements, including normal English words used in an absurd radiology "
    "context and random strings that do not follow a repeated-character pattern. "
    "Use category 'unrelated' only if the text is plainly extraneous or "
    "semantically incoherent in the given report context. Quote the SMALLEST "
    "exact offending token or phrase to avoid replacing a whole medical claim. "
    "If the intended replacement is unclear, give verification advice, not "
    "an invented replacement. Do not flag a genuine medical word merely "
    "because it is uncommon, misspelled, unfamiliar, or clinically unexpected; "
    "use language for supported spelling problems and logic for real internal "
    "medical contradictions. Names, signature blocks, doctor credentials, "
    "patient metadata, imaging protocols, devices, measurements, and abbreviations "
    "must not be flagged as unrelated just because they are non-narrative. "
    "Then proofread spelling, grammar, punctuation and awkward constructions. "
    "Only flag well-supported errors. Never invent imaging findings or diagnoses. "
    "Quotes MUST be exact substrings of supplied report and no longer than 500 "
    "characters. Use category logic only for potentially clinically significant "
    "internal contradictions, language for wording concerns, and unrelated "
    "for unmistakable contextual nonsense or off-topic fragments. "
    "Do not repeat patient identifiers in the explanation or suggestion. "
    "Return only JSON: {\"issues\":[{\"quote\":\"exact report text\","
    "\"category\":\"logic or language or unrelated\",\"explanation\":\"brief explanation\","
    "\"suggestion\":\"optional proposed wording\"}]}. "
    "Keep suggestions clinically conservative; if uncertain, recommend verification."
)


# A dedicated, independent semantic audit of every Conclusion assertion.
CONCLUSION_SYSTEM = (
    "You are a careful radiology Conclusion-versus-Findings consistency auditor. "
    "Read every assertion of the Conclusion/Impression and cross-check with Findings. "
    "Flag clinically meaningful contradictions, unrelated or nonsensical claims, "
    "wrong laterality, negation, organ, size, severity, disease, and any diagnosis "
    "or conclusion assertion not supported by the supplied Findings. "
    "Flag a potentially important finding omitted from the Conclusion only if it "
    "could materially affect clinical interpretation, and phrase omissions as "
    "questions to the reporting radiologist, not automatic diagnoses. "
    "Preserve uncertainty: possible, likely and questionable are not definitive. "
    "Preserve legitimate medical abbreviations and acronyms without assuming "
    "they are nonsense. Dates and year references including 2026 and 2027 "
    "must not be treated as incorrect solely due to their value. "
    "For unsupported conclusions request confirmation instead of inventing facts. "
    "Do not regard every missing incidental finding as a contradiction. "
    "For each issue anchor an EXACT substring of the CONCLUSION (not Findings) "
    "up to 500 characters, preferably an entire unique sentence. "
    "Use category 'unrelated' for high-confidence nonsense in Conclusion, "
    "including contextually absurd real words, accidental filler, random "
    "strings, non-medical statements, or off-topic material. Anchor the "
    "smallest exact offending fragment if feasible. Never infer that an "
    "unfamiliar valid medical term is gibberish. Do not use 'unrelated' for an unsupported "
    "but medically relevant statement; use 'logic' for that. "
    "Return only JSON with issues array; each item: quote, category='logic' or 'unrelated', "
    "explanation, suggestion. If nothing is wrong return an empty issues array. "
    "Never repeat patient identifiers."
)
# Word's Document.Content.Text uses CR-only paragraphs and 0x07 table-cell
# boundaries. ^ with re.MULTILINE only handles LF, silently skipping a
# Conclusion table in real Word reports. Recognize all three boundaries.
WORD_BOUNDARY = r"(?:^|(?<=[\r\n\x07]))"
CONCLUSION_HEADING = re.compile(
    WORD_BOUNDARY +
    r"[ \t]*\*{0,3}[ \t]*(?:conclusion|impression)[ \t]*\*{0,3}[ \t]*[:：]?[ \t]*",
    re.IGNORECASE
)
CONCLUSION_END = re.compile(
    WORD_BOUNDARY + r"[ \t]*(?:thanks[;,.]?|dr\.[ \t]|signed[ \t]*by)",
    re.IGNORECASE
)

def conclusion_section(report):
    match = CONCLUSION_HEADING.search(report)
    if not match:
        return None
    tail = report[match.start():]
    further = CONCLUSION_HEADING.search(tail, len(match.group(0)))
    signed = CONCLUSION_END.search(tail, len(match.group(0)))
    ends = [m.start() for m in (further, signed) if m]
    section = tail[:min(ends) if ends else len(tail)]
    return section if section.strip() else None

def conclusion_audit_context(report):
    section = conclusion_section(report)
    if not section:
        return None, None
    if len(report) <= 6000:
        return report, section
    if len(section)>5000:
        return None, None
    # Bounded context on very long reports; explicitly state partial coverage.
    findings = report[:report.find(section)]
    available = max(0, 5900-len(section)-90)
    return (findings[:available] +
            "\n[PARTIAL FINDINGS CONTEXT; CHECK FULL REPORT CLINICALLY]\n" +
            section), section


# Local high-precision safeguard for unmistakable biographical/employment
# statements inside clinical conclusions. It does not replace the semantic AI
# audit or make any medical claim. Never persist patient/report content.
OFF_TOPIC_EMPLOYMENT = re.compile(
    r"(?i)\b(?:work|works|working|employed)\s+(?:at|for)\s+\S+"
    r"|\b(?:my|his|her)\s+(?:workplace|employer|occupation)\b"
    r"|\b(?:is|am|are)\s+(?:a|an)\s+"
    r"(?:developer|accountant|engineer|teacher|student|employee)\b"
)

# Narrow repeated-motif detector; not a dictionary and not a medical
# spellchecker. Clinical abbreviations and unfamiliar real terms must never
# be silently rewritten. Only transparent physician-verify annotations.
REPEATED_GIBBERISH = re.compile(r"^([a-z]{2})\1+[a-z]?$")

def repeated_gibberish_tokens(line):
    return [m.group(0) for m in re.finditer(r"\b[A-Za-z]{5,12}\b", line)
            if REPEATED_GIBBERISH.fullmatch(m.group(0))]

def whole_report_gibberish_groups(chunks):
    """Find only repeated-motif nonsense; never dictionary-check clinical terms."""
    groups = []
    for part in chunks:
        seen, issues = set(), []
        for token in repeated_gibberish_tokens(part):
            if token in seen:
                continue
            seen.add(token)
            issues.append({
                "quote": token,
                "category": "unrelated",
                "explanation": "Unintelligible repeated-letter token in radiology report.",
                "suggestion": "Verify intended wording with reporting radiologist.",
                "section": "body",
            })
        groups.append(issues)
    return groups

def obvious_nonmedical_conclusion_issues(report):
    section = conclusion_section(report)
    if not section:
        return []
    # Word text uses CR paragraph breaks, ordinary text may use LF.
    issues = []
    for raw in re.split(r"[\r\n\x07]+", section):
        line = raw.strip().lstrip("o•◦- \t").strip()
        if not line or len(line) > 300:
            continue
        if OFF_TOPIC_EMPLOYMENT.search(line):
            issues.append({
                "quote": line,
                "category": "unrelated",
                "explanation": ("Non-medical employment or personal statement "
                                "appears in the clinical conclusion."),
                "suggestion": "Ask the reporting radiologist to verify and remove if unintended.",
                "section": "conclusion",
            })
        for token in repeated_gibberish_tokens(line):
            issues.append({
                "quote": token,
                "category": "unrelated",
                "explanation": "Unintelligible repeated-letter token in the clinical conclusion.",
                "suggestion": "Confirm the intended text with the reporting radiologist.",
                "section": "conclusion",
            })
    return issues

def distribute_conclusion_issues(chunks, issues):
    section = conclusion_section("".join(chunks))
    groups = [[] for _ in chunks]
    if not section:
        return groups
    seen = set()
    for issue in issues:
        quote = issue.get("quote", "")
        if not quote or quote not in section or issue.get("category") not in ("logic", "unrelated"):
            continue
        candidates = [i for i, chunk in enumerate(chunks) if quote in chunk]
        if not candidates:
            continue
        index = candidates[-1]
        ident = (index, quote, issue["category"])
        if ident in seen:
            continue
        seen.add(ident)
        tagged = dict(issue)
        tagged["section"] = "conclusion"
        groups[index].append(tagged)
    return groups

def conclusion_issue_groups(key, chunks, reviewer=None):
    full_report = "".join(chunks)
    context, section = conclusion_audit_context(full_report)
    groups = [[] for _ in chunks]
    if not context:
        return groups
    if reviewer is None:
        reviewer = lambda content: api_review(
            key, content, system_prompt=CONCLUSION_SYSTEM,
            user_prefix="Audit every assertion in the Conclusion against the Findings:\n\n")
    return distribute_conclusion_issues(chunks, reviewer(context))

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
        if quote not in text or category not in ("logic", "language", "unrelated"):
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
        if category not in ("logic", "language", "unrelated"):
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
        with open_url(req, timeout=70) as response:
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
        if code == 499:
            raise ReviewFailure("GROQ_CLIENT_DISCONNECTED", True)
        if code == 408 or code == 429:
            raise ReviewFailure("GROQ_HTTP_" + str(code), True, retry_after_seconds(exc.headers))
        if code == 503:
            raise ReviewFailure("GROQ_SERVER_OVER_CAPACITY", True, retry_after_seconds(exc.headers))
        if code in (500, 502, 504):
            raise ReviewFailure("GROQ_HTTP_" + str(code), True, retry_after_seconds(exc.headers))
        raise ReviewFailure("GROQ_HTTP_" + str(code))
    except (urllib.error.URLError, TimeoutError, ConnectionError):
        raise ReviewFailure("GROQ_NETWORK", True)


def api_review(key, text, requester=transport_request, sleeper=time.sleep,
               system_prompt=SYSTEM, user_prefix="Review this report completely:\n\n"):
    request_body = {
        "model": MODEL,
        "messages": [{"role": "system", "content": system_prompt},
                     {"role": "user", "content": user_prefix + text}],
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
                 "User-Agent": "Haronex-Radiology-Agent/11"})
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
        if version != "haronex-groq-qwen38-v11":
            raise ValueError("Unsupported review version")
        # Two independent reviews run concurrently rather than sequentially.
        # Both are still mandatory, and their results are combined only after
        # successful completion. This reduces wall time, not model token usage.
        with ThreadPoolExecutor(max_workers=2) as executor:
            general = executor.submit(
                lambda: [review_with_recovery(key, text) for text in texts])
            clinical = executor.submit(conclusion_issue_groups, key, texts)
            groups = general.result()
            independent_audit = clinical.result()
        local_offtopic = distribute_conclusion_issues(
            texts, obvious_nonmedical_conclusion_issues("".join(texts)))
        for i, added in enumerate(independent_audit):
            existing = {(x["quote"], x["category"]) for x in groups[i]}
            groups[i].extend(x for x in added
                             if (x["quote"], x["category"]) not in existing)
        # The local Conclusion audit is section-limited; this independent
        # safeguard also sees Findings, tables and standalone symbol-wrapped text.
        for i, added in enumerate(whole_report_gibberish_groups(texts)):
            existing = {(x["quote"], x["category"]) for x in groups[i]}
            groups[i].extend(x for x in added
                             if (x["quote"], x["category"]) not in existing)
        for i, added in enumerate(local_offtopic):
            existing = {(x["quote"], x["category"]) for x in groups[i]}
            groups[i].extend(x for x in added
                             if (x["quote"], x["category"]) not in existing)
        print(json.dumps({"ok": True, "mode": "groq", "passes": 2,
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
                      "GROQ_HTTP_502", "GROQ_HTTP_503", "GROQ_HTTP_504",
                      "GROQ_SERVER_OVER_CAPACITY", "GROQ_CLIENT_DISCONNECTED"}
        code = message if (message in safe_codes or message.startswith("GROQ_REQUEST_400_")) else "WORKER_" + type(exc).__name__.upper()
        print(json.dumps({"ok": False, "error": code}))

def self_test():
    sample = "Mild fatty hepatomegaly. Small fat containing umbilical hernia."
    issues = normalize(sample, [{"quote": "Small fat containing umbilical hernia",
                                "category": "language", "explanation": "Missing punctuation",
                                "suggestion": "Small fat-containing umbilical hernia."}])
    assert len(issues) == 1 and issues[0]["category"] == "language"
    clinical = ("Findings: Multiple bilateral renal cortical cysts, bulky prostate.\n"
                "Conclusion:\nNormal kidneys with no cysts.\nThanks;")
    section = conclusion_section(clinical)
    assert section and "Normal kidneys with no cysts." in section
    assert "Thanks" not in section
    fake = {"quote": "Normal kidneys with no cysts.", "category": "logic",
            "explanation": "Contradicts the documented renal cysts",
            "suggestion": "Reconcile with Findings after physician verification."}
    assert conclusion_issue_groups("offline", [clinical],
        reviewer=lambda context: [fake]) == [[dict(fake, section="conclusion")]]
    assert conclusion_issue_groups("offline", ["Normal report."],
        reviewer=lambda context: [fake]) == [[]]
    off_topic = clinical.replace("Normal kidneys with no cysts.", "Haron work at dokki scan")
    outside = {"quote": "Haron work at dokki scan", "category": "unrelated",
               "explanation": "Statement is not a radiology finding",
               "suggestion": "Verify and remove if entered inadvertently."}
    tagged = conclusion_issue_groups("offline", [off_topic], reviewer=lambda ctx: [outside])
    assert len(tagged[0]) == 1 and tagged[0][0]["category"] == "unrelated"
    assert normalize_response(off_topic, [outside])[0]["category"] == "unrelated"
    assert conclusion_section(
        "Findings: Bilateral cysts.\r\x07Conclusion:\r\x07Haron work at dokki scan"
        "\r\x07Thanks;") is not None
    assert obvious_nonmedical_conclusion_issues(
        "Findings: Bilateral cysts.\r\x07Conclusion:\r\x07"
        "o Haron work at dokki scan\r\x07Thanks;")
    assert obvious_nonmedical_conclusion_issues(
        "Findings: Bilateral renal cysts.\rConclusion:\r"
        "o Haron work at dokki scan\r"
        "o Bilateral renal cortical cysts\rThanks;")[0]["category"] == "unrelated"
    assert obvious_nonmedical_conclusion_issues(
        "Findings: Bulky prostate.\rConclusion:\rBulky prostate.\rThanks;") == []
    assert obvious_nonmedical_conclusion_issues(
        "Findings: The patient works at the hospital.\r"
        "Conclusion:\rNo acute finding.\rThanks;") == []
    # Both checks are mandatory while their execution overlaps.
    events = []
    def simulated(tag):
        events.append(tag)
        return tag
    with ThreadPoolExecutor(max_workers=2) as pool:
        a = pool.submit(simulated, "general")
        b = pool.submit(simulated, "conclusion")
        assert a.result() == "general" and b.result() == "conclusion"
    nonsense = ("Findings: Bilateral cysts.\rConclusion:\r"
                "Element of dsdsd urinary cystitis.\rThanks;")
    nonsense_issues = obvious_nonmedical_conclusion_issues(nonsense)
    assert len(nonsense_issues) == 1
    assert nonsense_issues[0]["quote"] == "dsdsd"
    assert nonsense_issues[0]["category"] == "unrelated"
    assert repeated_gibberish_tokens("Bilateral renal cortical cysts.") == []
    assert repeated_gibberish_tokens("MRI MRCP IHBR PETCT BMD 2026/2027") == []
    assert repeated_gibberish_tokens("dsdsd") == ["dsdsd"]
    assert repeated_gibberish_tokens("***sasasasas ***") == ["sasasasas"]
    assert whole_report_gibberish_groups(["Findings: sasasasas"]) == [[{
        "quote": "sasasasas", "category": "unrelated",
        "explanation": "Unintelligible repeated-letter token in radiology report.",
        "suggestion": "Verify intended wording with reporting radiologist.",
        "section": "body",
    }]]
    assert [i["quote"] for i in whole_report_gibberish_groups(
        ["Findings: ***dsdsd ***", "Conclusion:\\rElement of dsdsd urinary cystitis."])[0]] == ["dsdsd"]
    assert whole_report_gibberish_groups(["CT, MRCP, IHBR, MRI, 2026/2027."]) == [[]]
    # Regression contract: model checks full-report semantics, not just regex.
    assert "EVERY clinical sentence" in SYSTEM
    assert "normal English words used in an absurd radiology" in SYSTEM
    assert "category 'unrelated'" in SYSTEM
    assert "contextually absurd real words" in CONCLUSION_SYSTEM
    assert "genuine medical word merely" in SYSTEM
    assert MODEL == "qwen/qwen3.8-27b"
    class FakeError:
        def read(self, limit):
            return b'{"error":{"message":"unsupported reasoning effort"}}'
    assert classify_http_400(FakeError()) == "GROQ_REQUEST_400_REASONING"
    assert "Haronex-Radiology-Agent/10.19" in str(api_review.__code__.co_consts)
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
    assert len(called) == 3 and delays == [2, 6]
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
        assert len(called) == 4 and delays == [4, 6, 15]
    assert retry_after_seconds({"Retry-After": "999"}) == 30
    from io import BytesIO
    class Http503(urllib.error.HTTPError):
        def __init__(self):
            super().__init__(ENDPOINT, 503, "Service unavailable", {}, BytesIO(b'{"error":{"message":"over capacity"}}'))
    def capacity_url(req, timeout=70):
        assert timeout == 70
        raise Http503()
    try:
        transport_request(None, capacity_url)
        raise AssertionError("503 treated as successful")
    except ReviewFailure as exc:
        assert exc.code == "GROQ_SERVER_OVER_CAPACITY" and exc.retryable
    class Http499(urllib.error.HTTPError):
        def __init__(self):
            super().__init__(ENDPOINT, 499, "Client disconnected", {}, BytesIO(b'{}'))
    try:
        transport_request(None, lambda req, timeout=70: (_ for _ in ()).throw(Http499()))
        raise AssertionError("499 treated as success")
    except ReviewFailure as exc:
        assert exc.code == "GROQ_CLIENT_DISCONNECTED" and exc.retryable
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
