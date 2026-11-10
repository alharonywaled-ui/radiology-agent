# Haronex experimental Gemini Web browser worker. No API keys or passwords.
# Browser must be signed out; disposable Guest mode profile, no account/session reuse.
# Protocol: one JSON object on stdin; one JSON object on stdout.
import json
import os
import sys
import time
import tempfile
from pathlib import Path

URL = "https://gemini.google.com/app"
SELECTORS = (
    ".model-response-text",
    "message-content",
    "[data-message-author-role='model']",
    ".response-content",
)

def result(ok, **kwargs):
    sys.stdout.write(json.dumps(dict(ok=ok, **kwargs), ensure_ascii=False) + "\n")
    sys.stdout.flush()

def balanced_json(source):
    pos = source.find("{")
    while pos >= 0:
        depth, quote, escape = 0, False, False
        for index in range(pos, len(source)):
            char = source[index]
            if quote:
                if escape:
                    escape = False
                elif char == "\\":
                    escape = True
                elif char == '"':
                    quote = False
            else:
                if char == '"':
                    quote = True
                elif char == "{":
                    depth += 1
                elif char == "}":
                    depth -= 1
                    if depth == 0:
                        try:
                            parsed = json.loads(source[pos:index + 1])
                        except (ValueError, TypeError):
                            break
                        if isinstance(parsed, dict) and isinstance(parsed.get("issues"), list):
                            return parsed
                        break
        pos = source.find("{", pos + 1)
    return None

def response_texts(driver):
    return driver.execute_script("""
        const selectors = arguments[0];
        const nodes = new Set();
        for(const selector of selectors) {
            try {document.querySelectorAll(selector).forEach(n => nodes.add(n));}catch(e){}
        }
        return Array.from(nodes).map(n => (n.innerText || '').trim()).filter(Boolean);
    """, SELECTORS) or []

def editor(driver):
    for css in ('rich-textarea [contenteditable="true"]',
                'div.ql-editor[contenteditable="true"]',
                '[role="textbox"][contenteditable="true"]',
                '[contenteditable="true"]'):
        matches = driver.find_elements("css selector", css)
        for item in matches:
            if item.is_displayed() and item.is_enabled():
                return item
    return None

def send_prompt(driver, prompt):
    from selenium.webdriver.common.keys import Keys
    field = editor(driver)
    if field is None:
        raise RuntimeError("Gemini composer unavailable. Sign in or check site layout.")
    field.click()
    # CDP inserts full UTF-8 text in one operation, avoiding slow per-character keystrokes.
    try:
        driver.execute_cdp_cmd("Input.insertText", {"text": prompt})
    except Exception:
        field.send_keys(prompt)
    time.sleep(0.2)
    for selector in ('button[aria-label*="Send"]', 'button[aria-label*="send"]',
                     'button[data-test-id*="send"]', 'button.send-button'):
        for button in driver.find_elements("css selector", selector):
            if button.is_displayed() and button.is_enabled():
                button.click()
                return
    field.send_keys(Keys.ENTER)

def wait_for_response(driver, previous_count, seconds=160):
    deadline = time.monotonic() + seconds
    stable = None
    stable_since = None
    while time.monotonic() < deadline:
        texts = response_texts(driver)
        # On every review, a fresh conversation is opened. Do not accept old responses.
        candidates = texts[previous_count:] if len(texts) > previous_count else []
        for text in candidates:
            parsed = balanced_json(text)
            if parsed is None:
                continue
            marker = json.dumps(parsed, ensure_ascii=False, sort_keys=True)
            if marker != stable:
                stable = marker
                stable_since = time.monotonic()
            elif stable_since and time.monotonic() - stable_since >= 1.5:
                return parsed
        time.sleep(0.5)
    raise RuntimeError("No complete JSON answer received within the timeout. Gemini Web may have changed.")

def audit_prompt(report, prior_issues):
    # A genuinely second model call, not just rephrasing the first answer.
    # Do not assume the first answer is complete or correct.
    known = [
        {"quote": issue.get("quote", ""), "category": issue.get("category", "")}
        for issue in prior_issues if isinstance(issue, dict)
    ]
    return (
        "You are Haronex's independent second-pass audit reviewer. "
        "Read the ENTIRE report again, sentence by sentence. "
        "Search independently for errors missed by the first review, checking: "
        "(1) spelling, grammar, punctuation, sentence clarity; "
        "(2) internal consistency of clinical terminology and abbreviations; "
        "(3) contradictions between findings and impression, right versus left, "
        "negation, measurements, counts and anatomical descriptions; "
        "(4) any remaining distinct internally supported errors. "
        "Known findings to avoid duplicating (not guaranteed to be correct): "
        + json.dumps(known, ensure_ascii=False) +
        "\\nReport to audit:\\n" + report +
        "\\nReturn ONLY one JSON object: "
        '{"issues":[{"quote":"exact substring from report","category":"language or logic",'
        '"explanation":"brief evidence-based reason","suggestion":"optional wording"}]}. '
        "Each quote MUST appear exactly in the report. Include only NEW, well-supported "
        "issues not already listed. Do not invent medical diagnoses. "
        "An empty issues array is valid when no additional reliable issues exist. "
        "Clinical verification remains required."
    )


def normalize_issues(report, *groups):
    # Deterministic merge and conservative filtering. Order by source location.
    entries = {}
    for issues in groups:
        for item in issues:
            if not isinstance(item, dict):
                continue
            quote, category = item.get("quote"), item.get("category")
            if not isinstance(quote, str) or not 2 <= len(quote) <= 500:
                continue
            if not isinstance(category, str):
                continue
            category = category.strip().lower()
            if category not in ("language", "logic") or quote not in report:
                continue
            key = quote  # same range should not carry duplicate comments
            record = {
                "quote": quote, "category": category,
                "explanation": str(item.get("explanation") or "")[:500],
                "suggestion": str(item.get("suggestion") or "")[:250],
            }
            if key not in entries or (category == "logic" and entries[key]["category"] != "logic"):
                entries[key] = record
    return sorted(entries.values(), key=lambda row: (
        report.find(row["quote"]), row["category"], row["quote"]
    ))


def main():
    try:
        payload = json.loads(sys.stdin.read())
        prompts = payload.get("prompts", [])
        texts = payload.get("texts", [])
        review_version = payload.get("review_version", "")
        if not isinstance(prompts, list) or not prompts or len(prompts) > 20:
            raise ValueError("Invalid prompt list.")
        if not isinstance(texts, list) or len(texts) != len(prompts):
            raise ValueError("Missing matching original report parts.")
        if not isinstance(review_version, str) or not review_version.startswith("haronex-comprehensive-v5-"):
            raise ValueError("Unsupported review version.")
        if any(not isinstance(p, str) or not p or len(p) > 30000 for p in prompts):
            raise ValueError("One or more prompts are invalid.")
        if any(not isinstance(t, str) or not t or len(t) > 12000 for t in texts):
            raise ValueError("Invalid report parts.")
        from selenium import webdriver
        from selenium.webdriver.edge.options import Options
        # An isolated temporary profile is deliberately not the user's Edge profile.
        # Never use stored Google sign-in cookies or credentials.
        with tempfile.TemporaryDirectory(prefix="HaronexGuest_") as fresh_profile:
            options = Options()
            options.add_argument("--user-data-dir=" + fresh_profile)
            options.add_argument("--guest")
            options.add_argument("--headless=new")
            options.add_argument("--no-first-run")
            options.add_argument("--disable-notifications")
            options.add_argument("--window-size=1100,800")
            driver = webdriver.Edge(options=options)
            try:
                all_issues = []
                for index, prompt in enumerate(prompts):
                    collected = []
                    for pass_index in (0, 1):
                        driver.get(URL)
                        deadline = time.monotonic() + 25
                        while editor(driver) is None and time.monotonic() < deadline:
                            time.sleep(0.5)
                        if editor(driver) is None:
                            raise RuntimeError("Guest text input unavailable. Gemini may require sign-in or may be unsupported for this location/device.")
                        before = len(response_texts(driver))
                        content = prompt if pass_index == 0 else audit_prompt(texts[index], collected)
                        send_prompt(driver, content)
                        parsed = wait_for_response(driver, before)
                        if not isinstance(parsed, dict) or not isinstance(parsed.get("issues"), list):
                            raise RuntimeError("Incomplete response in comprehensive review.")
                        if pass_index == 0:
                            collected = normalize_issues(texts[index], parsed["issues"])
                        else:
                            collected = normalize_issues(texts[index], collected, parsed["issues"])
                    all_issues.append(collected)
                result(True, groups=all_issues, requests=len(prompts) * 2,
                       passes=2, mode="guest", review_version=review_version)
            finally:
                driver.quit()
    except Exception as exc:
        result(False, error=str(exc)[:400])

if __name__ == "__main__":
    if "--self-test" in sys.argv:
        sample = 'Some text before { "issues":[{"quote":"abc","category":"language"}] } after'
        parsed = balanced_json(sample)
        merged = normalize_issues("abc xyz", [
            {"quote":"abc","category":"language","explanation":"typo"},
            {"quote":"abc","category":"language","explanation":"typo"},
            {"quote":"missing","category":"logic"}
        ], [{"quote":"xyz","category":"logic","explanation":"conflict"}])
        second = audit_prompt("abc xyz", merged)
        ok = (parsed is not None and len(parsed.get("issues", [])) == 1
              and len(merged) == 2 and merged[0]["quote"] == "abc"
              and merged[1]["quote"] == "xyz" and "abc xyz" in second)
        result(ok, test="parser_and_comprehensive_merge")
    else:
        main()
