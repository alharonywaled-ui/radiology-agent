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

def api_review(key, text):
    request_body = {
        "model": MODEL,
        "messages": [{"role": "system", "content": SYSTEM},
                     {"role": "user", "content": "Review this report completely:\n\n" + text}],
        "response_format": {"type": "json_object"},
        "reasoning_effort": "low",
        "temperature": 0.2,
        "max_completion_tokens": 1600
    }
    data = json.dumps(request_body, ensure_ascii=False).encode("utf-8")
    req = urllib.request.Request(ENDPOINT, data=data, method="POST",
        headers={"Authorization": "Bearer " + key,
                 "Content-Type": "application/json",
                 "Accept": "application/json"})
    try:
        with urllib.request.urlopen(req, timeout=45) as response:
            raw = response.read(3 * 1024 * 1024)
    except urllib.error.HTTPError as exc:
        # Avoid logging response headers/body or access credentials.
        raise RuntimeError("Groq HTTP status " + str(exc.code))
    except urllib.error.URLError:
        raise RuntimeError("Unable to reach Groq API")
    obj = json.loads(raw.decode("utf-8"))
    choices = obj.get("choices", [])
    if not choices:
        raise ValueError("Groq returned no choices")
    choice = choices[0]
    if choice.get("finish_reason") not in ("stop",):
        raise ValueError("Groq response truncated or incomplete")
    message = choice.get("message") or {}
    content = message.get("content")
    if not isinstance(content, str) or not content.strip():
        raise ValueError("Groq returned no review JSON")
    parsed = json.loads(content)
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
        if version != "haronex-groq-qwen38-v7-1":
            raise ValueError("Unsupported review version")
        groups = [api_review(key, text) for text in texts]
        print(json.dumps({"ok": True, "mode": "groq", "passes": 1,
                          "review_version": version, "groups": groups},
                         ensure_ascii=False))
    except Exception as exc:
        # No error bodies, secrets or reports in stdout.
        print(json.dumps({"ok": False, "error": type(exc).__name__ + ": " + str(exc)[:120]}))

def self_test():
    sample = "Mild fatty hepatomegaly. Small fat containing umbilical hernia."
    issues = normalize(sample, [{"quote": "Small fat containing umbilical hernia",
                                "category": "language", "explanation": "Missing punctuation",
                                "suggestion": "Small fat-containing umbilical hernia."}])
    assert len(issues) == 1 and issues[0]["category"] == "language"
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
