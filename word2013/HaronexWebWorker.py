# Haronex experimental Gemini Web browser worker. No API keys or passwords.
# Only the user-controlled, logged-in browser session is used.
# Protocol: one JSON object on stdin; one JSON object on stdout.
import json
import os
import sys
import time
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
    time.sleep(0.9)
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
            elif stable_since and time.monotonic() - stable_since >= 3:
                return parsed
        time.sleep(1)
    raise RuntimeError("No complete JSON answer received within the timeout. Gemini Web may have changed.")

def main():
    try:
        payload = json.loads(sys.stdin.read())
        prompts = payload.get("prompts", [])
        if not isinstance(prompts, list) or not prompts or len(prompts) > 20:
            raise ValueError("Invalid prompt list.")
        if any(not isinstance(p, str) or not p or len(p) > 30000 for p in prompts):
            raise ValueError("One or more prompts are invalid.")
        from selenium import webdriver
        from selenium.webdriver.edge.options import Options
        base = Path(os.environ.get("LOCALAPPDATA", str(Path.home()))) / "Haronex" / "BrowserProfile"
        base.mkdir(parents=True, exist_ok=True)
        options = Options()
        options.add_argument("--user-data-dir=" + str(base))
        options.add_argument("--no-first-run")
        options.add_argument("--disable-notifications")
        options.add_argument("--window-size=1100,800")
        driver = webdriver.Edge(options=options)
        try:
            all_issues = []
            for index, prompt in enumerate(prompts):
                driver.get(URL)
                deadline = time.monotonic() + 120
                while editor(driver) is None and time.monotonic() < deadline:
                    time.sleep(2)
                if editor(driver) is None:
                    raise RuntimeError("Gemini Web login or verification required. Sign in once in the opened Edge window, then retry.")
                # Minimize while web service processes prompts. Never bypass login/verification.
                try:
                    driver.minimize_window()
                except Exception:
                    pass
                before = len(response_texts(driver))
                send_prompt(driver, prompt)
                parsed = wait_for_response(driver, before)
                all_issues.append(parsed["issues"])
            result(True, groups=all_issues, requests=len(prompts))
        finally:
            driver.quit()
    except Exception as exc:
        result(False, error=str(exc)[:400])

if __name__ == "__main__":
    main()
