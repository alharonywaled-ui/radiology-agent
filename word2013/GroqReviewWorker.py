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
from collections import OrderedDict
from threading import Condition, RLock
import hashlib
import io
import os
import urllib3

MODEL = "qwen/qwen3.8-27b"
RETRY_DELAYS = (2, 6, 15)  # Bounded exponential backoff, four total HTTP attempts.
MAX_COMPLETION_TOKENS = 8192
MIN_SPLIT_CHARS = 600
MAX_SPLIT_DEPTH = 3
FORMAT_RETRIES = 1
DISCARDED_ISSUES = {"count": 0}

ENDPOINT = "https://api.groq.com/openai/v1/chat/completions"

# Only performance counters are reported; never include PHI, prompt or key.
_METRICS_LOCK = RLock()
_METRICS = {}
def reset_metrics():
    with _METRICS_LOCK:
        _METRICS.clear()
        _METRICS.update(http_requests=0, cache_hits=0, retries=0,
                        rate_limits=0, backoff_ms=0, json_generation_retries=0,
                        suppressed_partial_logic=0,
                        supported_conclusion_claims=0,
                        unsupported_conclusion_candidates=0)
def metric(name, increment=1):
    with _METRICS_LOCK:
        _METRICS[name] = _METRICS.get(name, 0) + increment
def snapshot_metrics():
    with _METRICS_LOCK:
        return dict(_METRICS)

class AdaptiveLimiter:
    """Bounded live concurrency, with additive increase / multiplicative decrease.
    No extra/reduced review passes and no unbounded request creation.
    """
    def __init__(self):
        self.condition = Condition()
        self.limit = 4
        self.active = 0
        self.successes = 0
    def acquire(self):
        with self.condition:
            while self.active >= self.limit:
                self.condition.wait()
            self.active += 1
    def release(self, rate_limited=False, success=False):
        with self.condition:
            self.active = max(0, self.active - 1)
            if rate_limited:
                self.limit = max(1, self.limit // 2)
                self.successes = 0
            elif success:
                self.successes += 1
                if self.successes >= 10 and self.limit < 6:
                    self.limit += 1
                    self.successes = 0
            self.condition.notify_all()

LIMITER = AdaptiveLimiter()
# Thread-safe HTTP(S) keep-alive and connection pool. Never cache the API key.
HTTP_POOL = urllib3.PoolManager(num_pools=2, maxsize=6, block=True,
    retries=False, timeout=urllib3.Timeout(connect=12, read=70))

class ShortLivedSectionCache:
    """Only independent exact-text passes. Conclusion context ALWAYS re-runs.
    RAM-only, bounded to 96 items / 10min, hash-only keys, no disk artifacts.
    """
    def __init__(self):
        self.data = OrderedDict()
        self.lock = RLock()
        self.ttl = 600
        self.maximum = 96
    def get_or_run(self, pass_name, key, text, fn):
        identity = hashlib.sha256((
            "v11.19|" + MODEL + "|" + pass_name + "|" +
            hashlib.sha256(key.encode("utf-8")).hexdigest() + "|" +
            text).encode("utf-8")).digest()
        now = time.monotonic()
        with self.lock:
            for old_id, (expiry, _) in list(self.data.items()):
                if expiry <= now: self.data.pop(old_id, None)
            if identity in self.data:
                expiry, value = self.data.pop(identity)
                self.data[identity] = (expiry, value)
                metric("cache_hits")
                return [dict(x) for x in value]
        result = fn(text)
        with self.lock:
            self.data[identity] = (time.monotonic() + self.ttl,
                                   [dict(x) for x in result])
            while len(self.data) > self.maximum:
                self.data.popitem(last=False)
        return result
SECTION_CACHE = ShortLivedSectionCache()


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
    "If the supplied Findings is marked INCOMPLETE, it is only an EXCERPT. "
    "NEVER infer that a Conclusion claim is absent from the COMPLETE Findings "
    "because it is not mentioned in the currently supplied excerpt. "
    "For incomplete excerpts, flag only positive contradictions where the "
    "visible excerpt explicitly and directly disagrees with a Conclusion. "
    "Do not conflate inflammatory FDG activity with neoplastic FDG activity. "
    "Interpret 'any other' as excluding abnormalities already described. "
    "Ambiguous scope or wording is not an established internal contradiction. "
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
# Dedicated contextual audit: uncommon words are not errors by themselves.
# The report must make the word/phrase visibly non-medical or incoherent.
OFF_TOPIC_SYSTEM = (
    "You are a focused radiology report context auditor. Examine EVERY "
    "sentence in Findings, Conclusion and supporting report body. "
    "Identify unmistakable accidental filler, nonsensical words, or phrases "
    "whose meaning is completely incongruent with the specific medical "
    "sentence, even when they are real English words, not just repetitive "
    "keyboard patterns. Examples of the TYPE of error include an arbitrary "
    "token inserted between a preposition and a diagnosis, or an unrelated "
    "everyday activity described as an anatomical finding. "
    "Only report a fragment as 'unrelated' when the text itself strongly "
    "supports it. An unfamiliar medical term, recognized medical acronym, "
    "local abbreviation, patient name, physician signature, dates (including "
    "2026/2027), measurement, device, or anatomical variant is not automatically "
    "unrelated. Do not invent medical evidence. "
    "Return JSON ONLY in the form "
    "{\"issues\":[{\"quote\":\"exact shortest matching report fragment\","
    "\"category\":\"unrelated\",\"explanation\":\"why it is out of context\","
    "\"suggestion\":\"verification or safe literal correction\"}]}. "
    "Quote exact contiguous text, at most 500 characters. Never fabricate "
    "a replacement when intent is uncertain. Empty issues list is fine."
)

def contextual_unrelated_groups(key, chunks, reviewer=None):
    """Independent semantic audit for arbitrary off-topic language, every chunk."""
    if reviewer is None:
        reviewer = lambda chunk: api_review(
            key, chunk, system_prompt=OFF_TOPIC_SYSTEM,
            user_prefix="Audit this clinical text for ONLY clearly unrelated or "
                        "nonsensical material in context:\n\n")
    groups = []
    for chunk in chunks:
        found = reviewer(chunk)
        groups.append([dict(issue, section="body") for issue in found
                       if issue.get("category") == "unrelated"
                       and issue.get("quote") in chunk])
    return groups

# Separate exhaustive language-only audit. It runs on the SAME document
# snapshot as the semantic/Conclusion checks and the first general review.
LANGUAGE_SYSTEM = (
    "You are an expert English-language radiology copyeditor. Examine EACH "
    "sentence of the supplied report, including Findings, Side Findings, "
    "Conclusion, captions and tables. Identify clear grammar, spelling, "
    "punctuation, malformed word order, articles, and agreement errors. "
    "Do not skip a sentence because another error was found. "
    "Do not invent medical diagnoses, measurements, or findings. "
    "Preserve medical abbreviations (MRI, CT, MRCP, IHBR, BMD and local "
    "acronyms), physician names, identity numbers, measurements and valid "
    "dates including 2026 and 2027. "
    "Only flag high-confidence errors, category language only. "
    "For each issue quote the SMALLEST exact contiguous excerpt that fully "
    "identifies the error, and make suggestion be a LITERAL replacement for "
    "exactly that excerpt, not a full sentence for a two-word quote. "
    "If a whole sentence must change, quote the entire original sentence "
    "and give the entire replacement sentence. "
    'Return JSON only: {"issues":[{"quote":"exact text",'
    '"category":"language","explanation":"reason",'
    '"suggestion":"exact replacement"}]}. '
    "Return an empty array if no clear language errors exist."
)

def language_audit_groups(key, chunks, reviewer=None):
    if reviewer is None:
        reviewer=lambda part: review_with_recovery(
            key, part,
            review=lambda k,t: api_review(k,t,system_prompt=LANGUAGE_SYSTEM,
                        user_prefix="Perform a complete independent language audit:\n\n"))
    groups=[]
    for text in chunks:
        results=reviewer(text)
        groups.append([dict(item,section="body") for item in results
                       if item.get("category")=="language"
                       and item.get("quote") in text])
    return groups

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
    """Do not throw away findings-side contradictions just because the quote
    points to Findings rather than Conclusion. Preserve exact Word anchoring.
    """
    section = conclusion_section("".join(chunks))
    groups = [[] for _ in chunks]
    seen = set()
    for issue in issues:
        quote = issue.get("quote", "")
        if not quote or issue.get("category") not in ("logic", "unrelated"):
            continue
        candidates = [i for i, chunk in enumerate(chunks) if quote in chunk]
        if not candidates:
            continue
        in_conclusion = bool(section and quote in section)
        index = candidates[-1] if in_conclusion else candidates[0]
        ident = (index, quote, issue["category"])
        if ident in seen:
            continue
        seen.add(ident)
        tagged = dict(issue)
        tagged["section"] = "conclusion" if in_conclusion else "body"
        groups[index].append(tagged)
    return groups

def clinical_context_windows(full_report, max_chars=5900):
    """Lossless Findings coverage across every section of a long report.
    Re-run the independent clinical audit with the entire Conclusion in each
    bounded window. Never silently truncate Findings or skip clinical checks.
    """
    conclusion = conclusion_section(full_report)
    if not conclusion:
        raise ReviewFailure("GROQ_CLINICAL_CONCLUSION_NOT_FOUND")
    if len(full_report) <= max_chars:
        return [full_report]
    available = max_chars - len(conclusion) - 75
    if available < MIN_SPLIT_CHARS:
        raise ReviewFailure("GROQ_CLINICAL_CONTEXT_TOO_LONG")
    start = full_report.find(conclusion)
    findings = full_report[:start]
    # Include any clinical text after the Conclusion, before signing.
    post = full_report[start+len(conclusion):]
    source = findings + post
    windows = []
    for pos in range(0, len(source), available):
        piece = source[pos:pos+available]
        windows.append(
            "[INCOMPLETE FINDINGS EXCERPT — NOT THE FULL FINDINGS]\n"
            "The following excerpt is only one of several disjoint parts "
            "of the Findings. Missing detail in this excerpt is NOT proof "
            "that it is absent from the complete report. Never label an "
            "assertion unsupported solely because you cannot see its source "
            "here. Report only DIRECT, POSITIVE contradictions visible in "
            "this excerpt; ignore omissions and unresolved absence claims.\n"
            "Partial Findings/context excerpt:\n" + piece +
            "\nComplete Conclusion (compare against this excerpt only):\n" +
            conclusion)
    if not windows:
        raise ReviewFailure("GROQ_CLINICAL_CONTEXT_UNAVAILABLE")
    return windows

# Absence of evidence in ONE Findings window is not evidence of absence in
# the entire radiology report. Never represent it as a clinical contradiction.
_ABSENCE_INFERENCE = re.compile(
    r"(?i)(?:"
    r"(?:do(?:es)?\s+not|doesn['’]t|did\s+not|not)\s+"
    r"(?:contain|mention|describe|document|include|note|show|state|provide)"
    r"|(?:no|without)\s+(?:description|mention|documentation|evidence)\s+of"
    r"|(?:not|isn['’]t)\s+(?:described|documented|mentioned|included|provided)"
    r"|(?:unsupported|not\s+supported)\s+by\s+(?:the\s+)?"
    r"(?:available|provided|supplied|shown|partial)?\s*findings"
    r"|(?:absent|missing)\s+from\s+(?:the\s+)?"
    r"(?:provided|available|supplied|shown|partial)?\s*findings"
    r")"
)
def is_unverifiable_partial_logic(issue):
    if issue.get("category") != "logic":
        return False
    explanation = str(issue.get("explanation", ""))
    suggestion = str(issue.get("suggestion", ""))
    if _ABSENCE_INFERENCE.search(explanation) or _ABSENCE_INFERENCE.search(suggestion):
        return True
    # "Other" can refer to additional lesions, excluding ones already listed.
    # Positive inflammatory uptake in earlier Findings is not automatically
    # contradictory to "no OTHER active lesions". Require full context.
    quote = str(issue.get("quote", "")).lower()
    if ("rest of the body" in quote and "other" in quote and
        any(x in explanation.lower() for x in
            ("inflammat", "ground glass", "other", "neoplastic"))):
        return True
    return False


EVIDENCE_CHECK_SYSTEM = (
    "You are a radiology evidence locator, NOT a contradiction judge. "
    "Given a specific Conclusion assertion and ONE PARTIAL Findings excerpt, "
    "find any positive supporting evidence for that assertion in this excerpt. "
    "Paraphrases and equivalent clinical descriptions count as support. "
    "If evidence is present, return JSON {\"issues\":[{\"quote\":\"EXACT "
    "supporting excerpt from FINDINGS only\",\"category\":\"logic\","
    "\"explanation\":\"EVIDENCE_PRESENT\",\"suggestion\":\"\"}]}. "
    "If not present, return {\"issues\":[]}. Never infer absence from the "
    "excerpt, never invent findings and never copy text from Conclusion as evidence."
)
def clinical_evidence_in_window(key, conclusion_claim, findings_excerpt):
    # Limit input to one Findings window. Require exact source evidence quote.
    evidence = api_review(
        key, findings_excerpt,
        system_prompt=EVIDENCE_CHECK_SYSTEM,
        user_prefix="Conclusion claim to seek support for (NOT evidence):\n" +
                    conclusion_claim +
                    "\nSearch ONLY this Findings excerpt for matching support:\n")
    return any(row.get("quote","") in findings_excerpt and
               row.get("quote","").strip() and
               len(row["quote"]) >= 5 and
               row.get("explanation") == "EVIDENCE_PRESENT"
               for row in evidence)

def adjudicate_partial_absence(key, report, conclusion_claim,
                               evidence_reviewer=None):
    """Across ALL findings windows, suppress a missing-support claim if ANY
    piece gives supporting evidence. Lack of evidence in a single window is
    never enough. Fail closed on API errors, not false clinical success.
    """
    complete_conclusion = conclusion_section(report)
    if not complete_conclusion or conclusion_claim not in complete_conclusion:
        return False
    source_start=report.find(complete_conclusion)
    source=report[:source_start] + report[source_start+len(complete_conclusion):]
    # Reserve input for separate claim+instructions, split all evidence text.
    per_window=max(500, min(4000, 5100-len(conclusion_claim)))
    for pos in range(0,len(source),per_window):
        excerpt=source[pos:pos+per_window]
        if evidence_reviewer is None:
            found=clinical_evidence_in_window(key, conclusion_claim, excerpt)
        else:
            found=evidence_reviewer(conclusion_claim, excerpt)
        if found:
            metric("supported_conclusion_claims")
            return False
    metric("unsupported_conclusion_candidates")
    return True

def conclusion_issue_groups(key, chunks, reviewer=None, evidence_reviewer=None):
    full_report = "".join(chunks)
    contexts = clinical_context_windows(full_report)
    if reviewer is None:
        reviewer = lambda content: api_review(
            key, content, system_prompt=CONCLUSION_SYSTEM,
            user_prefix="Audit every assertion in the Conclusion against the Findings:\n\n")
    issues, possible_absences = [], []
    is_partial = len(full_report) > 5900
    for context in contexts:
        for item in reviewer(context):
            if is_partial and is_unverifiable_partial_logic(item):
                quote=item.get("quote","")
                # Scope-dependent ambiguous "other" phrases are wording
                # questions, not verified contradictions.
                if ("rest of the body" in quote.lower() and
                    "other" in quote.lower()):
                    metric("suppressed_partial_logic")
                    continue
                possible_absences.append(item)
            else:
                issues.append(item)
    # Check EACH questioned Conclusion assertion against every Findings
    # excerpt, even if a different audit window did not see its evidence.
    # Never silently discard genuine missing claims from the full report.
    for item in {row.get("quote",""):row for row in possible_absences
                 if row.get("quote","")}.values():
        if adjudicate_partial_absence(key,full_report,item["quote"],
                                      evidence_reviewer):
            verified=dict(item)
            verified["explanation"]=(
                "No supporting evidence was identified across the full "
                "Findings review; confirm with the reporting radiologist.")
            verified["suggestion"]=(
                "Verify or document supporting Findings for this "
                "Conclusion assertion.")
            issues.append(verified)
        else:
            metric("suppressed_partial_logic")
    return distribute_conclusion_issues(chunks, issues)

# Patient/organ cross-check uses EXPLICIT recorded sex ONLY. Names, titles
# and perceived identity are not reliable evidence of anatomical inventory.
_RECORDED_SEX=re.compile(
    r"(?i)\b(?:patient\s*(?:sex|gender)|sex|gender)\s*[:=\-]\s*"
    r"(male|female|m|f)\b|\b(male|female)\s+patient\b"
)
_FINDINGS_START=re.compile(r"(?i)\bfindings\s*[:：\-]")
_MALE_SPECIFIC=re.compile(r"(?i)\b(?:prostate|prostatic|testis|testes|"
                          r"testicle|testicular|seminal\s+vesicles?)\b")
_FEMALE_SPECIFIC=re.compile(r"(?i)\b(?:uterus|uterine|ovary|ovaries|"
                            r"ovarian|endometrium|endometrial|cervix|"
                            r"pregnancy|pregnant)\b")
_NEGATED_ANATOMY=re.compile(
    r"(?i)\b(?:no|absent|without|prior\s+removal\s+of|"
    r"history\s+of|status\s+post|post[\s-]operative)\b"
)
def patient_anatomy_consistency_groups(chunks):
    """Conservative *verification* flag, never an automatic sex determination.
    Avoid using names or assuming gender defines anatomy. Scan only Findings,
    not family history, patient metadata or differential context.
    """
    full="".join(chunks)
    start_match=_FINDINGS_START.search(full)
    if not start_match:
        return [[] for _ in chunks]
    header=full[:start_match.start()][:1600]
    sexes=set()
    for m in _RECORDED_SEX.finditer(header):
        found=(m.group(1) or m.group(2)).lower()
        sexes.add("male" if found in ("male","m") else "female")
    if len(sexes)!=1:
        return [[] for _ in chunks]
    sex=next(iter(sexes))
    conclusion=conclusion_section(full)
    body_end=full.find(conclusion) if conclusion else len(full)
    findings=full[start_match.end():body_end]
    matcher=_FEMALE_SPECIFIC if sex=="male" else _MALE_SPECIFIC
    groups=[[] for _ in chunks]
    offsets=[]
    position=0
    for chunk in chunks:
        offsets.append((position,position+len(chunk)))
        position+=len(chunk)
    used=set()
    for m in matcher.finditer(findings):
        before=findings[max(0,m.start()-90):m.start()]
        # Explicit negation, surgical history or quoted history is not a
        # current contradiction. Do not over-interpret uncertain anatomy.
        line=before.split("\r")[-1].split("\n")[-1]
        after=findings[m.end():m.end()+80]
        if (_NEGATED_ANATOMY.search(line) or
            re.search(r"(?i)\b(?:family\s+history|family\s+member|"
                      r"relative|mother|father|sister|brother)\b",after)):
            continue
        term=m.group(0)
        if term.lower() in used:
            continue
        used.add(term.lower())
        at=start_match.end()+m.start()
        for i,(left,right) in enumerate(offsets):
            if left<=at<right and at+len(term)<=right:
                groups[i].append({
                  "quote":term,
                  "category":"logic",
                  "explanation":(
                    "The report explicitly records sex as "+sex+
                    " while Findings describe "+term+
                    ". This is a possible metadata–anatomy mismatch, "
                    "not a definitive medical impossibility. Verify "
                    "recorded sex, anatomy, treatment and surgical history."),
                  "suggestion":"Verify patient information and anatomy "
                               "with the reporting clinician; do not "
                               "automatically alter this report.",
                  "section":"body"})
                break
    return groups


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
    # A model-side JSON generation failure is not a malformed request.
    # The service may return invalid_request_error as its generic "type";
    # detect the specific generation failure before that catch-all.
    if ("failed_generation" in error or
        any(token in diagnostic for token in (
        "json_validate_failed", "json validation failed",
        "failed_generation", "failed to generate valid json",
        "failed to produce valid json", "json_mode_failed"))):
        return "GROQ_REQUEST_400_JSON_GENERATION"
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
    """Persistent HTTPS pool for production, injectable urllib path for tests."""
    LIMITER.acquire()
    successful = False
    limited = False
    try:
        metric("http_requests")
        try:
            if open_url is urllib.request.urlopen:
                response = HTTP_POOL.request(
                    "POST", ENDPOINT, body=req.data,
                    headers=dict(req.header_items()),
                    timeout=urllib3.Timeout(connect=12, read=70),
                    retries=False, preload_content=True)
                status = response.status
                raw = response.data
                headers = response.headers
                response.release_conn()
                if status >= 400:
                    raise urllib.error.HTTPError(
                        ENDPOINT, status, "Groq request failed",
                        headers, io.BytesIO(raw))
            else:
                with open_url(req, timeout=70) as response:
                    raw = response.read(3 * 1024 * 1024)
            if len(raw) > 3 * 1024 * 1024:
                raise ReviewFailure("GROQ_INVALID_ENVELOPE")
            successful = True
            return raw
        except urllib.error.HTTPError as exc:
            code = exc.code
            if code == 400:
                safe_code = classify_http_400(exc)
                raise ReviewFailure(
                    safe_code,
                    retryable=(safe_code == "GROQ_REQUEST_400_JSON_GENERATION"))
            if code == 401:
                raise ReviewFailure("GROQ_AUTH_401")
            if code == 403:
                raise ReviewFailure("GROQ_FORBIDDEN_403")
            if code == 404:
                raise ReviewFailure("GROQ_MODEL_404")
            if code == 499:
                raise ReviewFailure("GROQ_CLIENT_DISCONNECTED", True)
            if code == 408 or code == 429:
                limited = code == 429
                if limited: metric("rate_limits")
                raise ReviewFailure("GROQ_HTTP_" + str(code), True,
                                    retry_after_seconds(exc.headers))
            if code == 503:
                limited = True
                metric("rate_limits")
                raise ReviewFailure("GROQ_SERVER_OVER_CAPACITY", True,
                                    retry_after_seconds(exc.headers))
            if code in (500, 502, 504):
                raise ReviewFailure("GROQ_HTTP_" + str(code), True,
                                    retry_after_seconds(exc.headers))
            raise ReviewFailure("GROQ_HTTP_" + str(code))
        except (urllib3.exceptions.HTTPError, urllib.error.URLError,
                TimeoutError, ConnectionError, OSError):
            raise ReviewFailure("GROQ_NETWORK", True)
    finally:
        LIMITER.release(rate_limited=limited, success=successful)


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
    # An explicit model-side JSON generation error may get ONE identical retry;
    # malformed HTTP 400 parameters, auth and safety failures never retry.
    recoverable_formats = {
        "GROQ_INVALID_ENVELOPE", "GROQ_INVALID_JSON",
        "GROQ_INVALID_STRUCTURE", "GROQ_ISSUES_NOT_ARRAY",
        "GROQ_INVALID_QUOTE", "GROQ_INVALID_CATEGORY",
        "GROQ_INVALID_WORDING", "GROQ_INVALID_ITEM",
        "GROQ_ALL_ISSUES_INVALID", "GROQ_EMPTY_RESPONSE",
        "GROQ_NO_CHOICES"
    }
    json_generation_retry_used = False
    for format_attempt in range(FORMAT_RETRIES + 1):
        for attempt in range(len(RETRY_DELAYS) + 1):
            try:
                raw = requester(req)
                break
            except ReviewFailure as exc:
                if exc.code == "GROQ_REQUEST_400_JSON_GENERATION":
                    if json_generation_retry_used or attempt >= len(RETRY_DELAYS):
                        raise
                    json_generation_retry_used = True
                    metric("retries")
                    metric("json_generation_retries")
                    metric("backoff_ms", 1000)
                    sleeper(1)
                    continue
                if not exc.retryable or attempt >= len(RETRY_DELAYS):
                    raise
                delay = RETRY_DELAYS[attempt]
                if exc.retry_after is not None:
                    delay = max(delay, exc.retry_after)
                metric("retries")
                metric("backoff_ms", int(delay * 1000))
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

def full_review_concurrently(key, texts, general_review=None,
                             conclusion_review=None, unrelated_review=None,
                             language_review=None):
    """Execute exactly the same four passes on an immutable snapshot.

    Every pass is mandatory; a failed future aborts the complete review.
    Results are assembled in original chunk order regardless of completion
    order, so Word quote anchoring and merge policy remain unchanged.
    """
    if not texts or len(texts) > 40:
        raise ValueError("Invalid report parts")
    if general_review is None:
        general_review = lambda text: SECTION_CACHE.get_or_run(
            "general", key, text, lambda part: review_with_recovery(key, part))
    if conclusion_review is None:
        conclusion_review = lambda parts: conclusion_issue_groups(key, parts)
    if unrelated_review is None:
        unrelated_review = lambda text: SECTION_CACHE.get_or_run(
            "unrelated", key, text,
            lambda part: contextual_unrelated_groups(key, [part])[0])
    if language_review is None:
        language_review = lambda text: SECTION_CACHE.get_or_run(
            "language", key, text,
            lambda part: language_audit_groups(key, [part])[0])
    max_workers = min(6, max(4, len(texts) * 2))
    with ThreadPoolExecutor(max_workers=max_workers) as pool:
        general = [pool.submit(general_review, t) for t in texts]
        conclusion = pool.submit(conclusion_review, texts)
        unrelated = [pool.submit(unrelated_review, t) for t in texts]
        language = [pool.submit(language_review, t) for t in texts]
        # Reading in index order does not block other submitted work.
        general_groups = [f.result() for f in general]
        clinical_groups = conclusion.result()
        unrelated_groups = [f.result() for f in unrelated]
        language_groups = [f.result() for f in language]
    if not (len(general_groups) == len(clinical_groups) ==
            len(unrelated_groups) == len(language_groups) == len(texts)):
        raise ValueError("Audit pass count mismatch")
    return general_groups, clinical_groups, unrelated_groups, language_groups


def main(raw_request=None):
    started=time.monotonic()
    reset_metrics()
    DISCARDED_ISSUES["count"] = 0
    try:
        incoming = json.loads(raw_request if raw_request is not None else sys.stdin.read())
        key = incoming.get("api_key")
        texts = incoming.get("texts")
        version = incoming.get("review_version")
        if not isinstance(key, str) or not key.startswith("gsk_") or len(key) < 20 or len(key) > 256:
            raise ValueError("Missing or invalid Groq API key")
        if not isinstance(texts, list) or not 1 <= len(texts) <= 40:
            raise ValueError("Invalid report parts")
        if any(not isinstance(x, str) or not x or len(x) > 6000 for x in texts):
            raise ValueError("Invalid report text")
        if version != "haronex-groq-qwen38-v11-19":
            raise ValueError("Unsupported review version")
        # Preserve ALL four independent audit passes; parallelize each report
        # chunk inside each pass instead of processing chunks serially. Bounded
        # worker count prevents runaway network threads on 40-part reports.
        groups, independent_audit, unrelated_audit, extra_language = (
            full_review_concurrently(key, texts))
        local_offtopic = distribute_conclusion_issues(
            texts, obvious_nonmedical_conclusion_issues("".join(texts)))
        for i, added in enumerate(independent_audit):
            existing = {(x["quote"], x["category"]) for x in groups[i]}
            groups[i].extend(x for x in added
                             if (x["quote"], x["category"]) not in existing)
        for i, added in enumerate(extra_language):
            existing = {(x["quote"], x["category"]) for x in groups[i]}
            groups[i].extend(x for x in added
                             if (x["quote"], x["category"]) not in existing)
        for i, added in enumerate(unrelated_audit):
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
        # Patient data vs anatomy: independent deterministic safety check,
        # only using an explicit recorded sex, never a patient name.
        for i, added in enumerate(patient_anatomy_consistency_groups(texts)):
            existing = {(x["quote"], x["category"]) for x in groups[i]}
            groups[i].extend(x for x in added
                             if (x["quote"], x["category"]) not in existing)
        # Reconcile ALL chunk-level "not found in Findings" observations
        # against the FULL report. True unsupported Conclusion assertions stay
        # visible, but partial-window omissions cannot become false warnings.
        full_report="".join(texts)
        if len(full_report)>5900:
            conclusion=conclusion_section(full_report) or ""
            verified_absences={}
            for i, group in enumerate(groups):
                safe_group=[]
                for item in group:
                    if not is_unverifiable_partial_logic(item):
                        safe_group.append(item)
                        continue
                    claim=item.get("quote","")
                    if "rest of the body" in claim.lower() and "other" in claim.lower():
                        metric("suppressed_partial_logic")
                        continue
                    if not claim or claim not in conclusion:
                        metric("suppressed_partial_logic")
                        continue
                    if claim not in verified_absences:
                        verified_absences[claim]=adjudicate_partial_absence(
                            key,full_report,claim)
                    if verified_absences[claim]:
                        amended=dict(item)
                        amended["explanation"]=(
                            "No supporting evidence was identified after "
                            "checking all Findings excerpts. The reporting "
                            "radiologist should verify this Conclusion assertion.")
                        amended["suggestion"]=(
                            "Confirm the finding, or document supporting "
                            "Findings before changing the Conclusion.")
                        safe_group.append(amended)
                    else:
                        metric("suppressed_partial_logic")
                groups[i]=safe_group
        print(json.dumps({"ok": True, "mode": "groq", "passes": 2,
                          "review_version": version, "groups": groups,
                          "discarded_issues": DISCARDED_ISSUES["count"],
                          "performance": dict(snapshot_metrics(),
                            elapsed_ms=int((time.monotonic()-started)*1000),
                            concurrency=LIMITER.limit)},
                         ensure_ascii=True))
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
                      "GROQ_SERVER_OVER_CAPACITY", "GROQ_CLIENT_DISCONNECTED",
                      "GROQ_CLINICAL_CONCLUSION_NOT_FOUND",
                      "GROQ_CLINICAL_CONTEXT_TOO_LONG",
                      "GROQ_CLINICAL_CONTEXT_UNAVAILABLE"}
        code = message if (message in safe_codes or message.startswith("GROQ_REQUEST_400_")) else "WORKER_" + type(exc).__name__.upper()
        print(json.dumps({"ok": False, "error": code,
                          "performance": dict(snapshot_metrics(),
                            elapsed_ms=int((time.monotonic()-started)*1000))}))

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
    try:
        conclusion_issue_groups("offline", ["Normal report."],
                                reviewer=lambda context: [fake])
        raise AssertionError("Missing Conclusion silently bypassed logic audit")
    except ReviewFailure as exc:
        assert exc.code == "GROQ_CLINICAL_CONCLUSION_NOT_FOUND"
    # Logic anchored to Findings MUST survive the dedicated audit.
    findings_quote = {"quote": "bilateral renal cortical cysts",
                      "category": "logic",
                      "explanation": "Contradicts normal kidney Conclusion",
                      "suggestion": "Verify conflicting kidney status"}
    assert conclusion_issue_groups("offline", [clinical],
          reviewer=lambda context: [findings_quote])[0][0]["section"] == "body"
    # Full Findings text must be covered; no first-5900-chars-only shortcut.
    long_report = ("Findings:\r" + "Normal interim finding.\r" * 400 +
                   "Unique contradiction near end of findings.\r" +
                   "Conclusion:\rNo abnormal findings.\r")
    windows = clinical_context_windows(long_report)
    assert len(windows)>1
    assert "Unique contradiction near end" in windows[-1]
    assert all("No abnormal findings" in w for w in windows)
    all_seen = []
    def audit_window(value):
        all_seen.append(value)
        if "Unique contradiction near end" in value:
            return [{"quote": "Unique contradiction near end of findings.",
                     "category": "logic", "explanation": "Conflicts with Conclusion",
                     "suggestion": "Verify discrepancy"}]
        return []
    combined = conclusion_issue_groups("offline",[long_report],reviewer=audit_window)
    assert len(all_seen)==len(windows) and combined[0][0]["category"]=="logic"
    # Regression: a complete report contains supporting lung/prostate Findings,
    # but each isolated context window may omit them. The AI must not convert
    # missing evidence in an EXCERPT into a false contradiction.
    false_claims = [
        ("The provided Findings do not contain any description of the left "
         "upper lobe mass.", "left upper lobe mass"),
        ("The available Findings do not mention the right pulmonary nodule.",
         "right pulmonary nodule"),
        ("The provided Findings do not describe the prostate size or uptake.",
         "enlarged prostate"),
    ]
    for explanation, _ in false_claims:
        assert is_unverifiable_partial_logic({
            "category": "logic", "quote": "Stable lesion",
            "explanation": explanation, "suggestion": "Verify section was included"})
    assert is_unverifiable_partial_logic({
        "category": "logic",
        "quote": "The rest of the body is free from any other metabolically active lesions.",
        "explanation": "The inflammatory FDG activity makes this ambiguous.",
        "suggestion": "Consider neoplastic lesions."})
    # A direct positive conflict must still be reported.
    assert not is_unverifiable_partial_logic({
        "category": "logic", "quote": "Normal kidneys with no cysts.",
        "explanation": "Findings explicitly describe bilateral renal cortical cysts.",
        "suggestion": "Verify contradictory statements."})
    assert not is_unverifiable_partial_logic({
        "category": "language", "quote": "Average sized",
        "explanation": "Consider hyphenation.", "suggestion": "Average-sized"})
    source_findings = (
        "Findings:\r"
        "Left upper lobe mass 2.3 x 2.9 cm, mildly FDG avid.\r" +
        "Other unrelated findings.\r" * 400 +
        "Right lower lobe nodule without FDG uptake.\r"
        "Markedly enlarged heterogeneous prostate with nodules.\r"
    )
    source_conclusion = (
        "Conclusion:\r"
        "Mild regression of left upper lobe mass.\r"
        "Sustained remission of right pulmonary nodule.\r"
        "Enlarged prostate with metabolically active nodules.\r"
        "The rest of the body is free from any other metabolically active lesions.\r"
    )
    full_source = source_findings + source_conclusion
    assert len(full_source) > 5900
    def false_window_audit(window):
        assert "[INCOMPLETE FINDINGS EXCERPT" in window
        return [
            {"quote": "Mild regression of left upper lobe mass.",
             "category": "logic", "explanation": false_claims[0][0],
             "suggestion": "Confirm full Findings included."},
            {"quote": "Sustained remission of right pulmonary nodule.",
             "category": "logic", "explanation": false_claims[1][0],
             "suggestion": "Confirm full Findings included."},
            {"quote": "Enlarged prostate with metabolically active nodules.",
             "category": "logic", "explanation": false_claims[2][0],
             "suggestion": "Confirm full Findings included."},
            {"quote": "The rest of the body is free from any other metabolically active lesions.",
             "category": "logic",
             "explanation": "Inflammatory ground glass uptake is described.",
             "suggestion": "Consider any other neoplastic lesions."},
        ]
    reset_metrics()
    evidence_terms={
        "Mild regression of left upper lobe mass.":"Left upper lobe mass",
        "Sustained remission of right pulmonary nodule.":"Right lower lobe nodule",
        "Enlarged prostate with metabolically active nodules.":"enlarged heterogeneous prostate"
    }
    checked=[]
    def located_evidence(claim, excerpt):
        checked.append((claim,excerpt))
        term=evidence_terms.get(claim)
        return bool(term and term.lower() in excerpt.lower())
    protected = conclusion_issue_groups(
        "offline", [full_source], reviewer=false_window_audit,
        evidence_reviewer=located_evidence)
    assert protected == [[]]
    assert checked and snapshot_metrics()["suppressed_partial_logic"] >= 4
    # A genuinely absent assertion must remain visible after *all* windows
    # were checked. Never treat a partial window as the entire Findings.
    absent_claim="Sustained remission of right pulmonary nodule."
    absent_report=full_source.replace(
        "Right lower lobe nodule without FDG uptake.\r","")
    def one_missing(window):
        return [{"quote":absent_claim,"category":"logic",
                 "explanation":"The provided Findings do not mention a "
                               "right pulmonary nodule.",
                 "suggestion":"Confirm full Findings were provided."}]
    actual_missing=conclusion_issue_groups(
        "offline",[absent_report],reviewer=one_missing,
        evidence_reviewer=lambda claim, excerpt: False)
    assert len(actual_missing[0])==1
    assert "full Findings review" in actual_missing[0][0]["explanation"]
    # Only *explicit* metadata can trigger possible patient–anatomy mismatch.
    male_report="Patient Name: Alex Smith\rMale patient 70 years old.\rFindings:\rUterine mass.\rConclusion:\rMass.\r"
    female_report="Sex: Female\rFindings:\rMarkedly enlarged prostate.\rConclusion:\rNeeds clinical review.\r"
    assert patient_anatomy_consistency_groups([male_report])[0][0]["category"]=="logic"
    assert patient_anatomy_consistency_groups([female_report])[0][0]["quote"]=="prostate"
    assert patient_anatomy_consistency_groups([
        "Patient Name: Maria\rFindings:\rProstate is enlarged.\rConclusion:\rFollow-up."])==[[]]
    assert patient_anatomy_consistency_groups([
        "Sex: Male\rFindings:\rNo uterus is identified.\rConclusion:\rNormal."])==[[]]
    assert patient_anatomy_consistency_groups([
        "Sex: Female\rFindings:\rProstate cancer in family history.\rConclusion:\rNormal."])==[[]]
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
    synthetic = "Findings: Element of essa urinary cystitis."
    stub = lambda text: [{"quote": "essa", "category": "unrelated",
                          "explanation": "Unrelated in context",
                          "suggestion": "Verify intended wording."}]
    assert contextual_unrelated_groups("offline", [synthetic], reviewer=stub) == [[
        {"quote": "essa", "category": "unrelated", "explanation": "Unrelated in context",
         "suggestion": "Verify intended wording.", "section": "body"}]]
    assert "2026/2027" in OFF_TOPIC_SYSTEM
    assert "EVERY" in OFF_TOPIC_SYSTEM
    mock_language=lambda text: [{
        "quote":"Average sized", "category":"language",
        "explanation":"Word choice", "suggestion":"Average-sized"}]
    assert language_audit_groups("offline", ["Average sized kidneys."],
                                 reviewer=mock_language)[0][0]["quote"]=="Average sized"
    assert "quote the entire original sentence" in LANGUAGE_SYSTEM
    # Four-pass concurrency must preserve every chunk, category and order.
    samples = ["Findings: Average sized kidneys.\r",
               "Conclusion: No urolithiasis.\r"]
    general_mock = lambda text: [{"quote": text[0:8], "category": "language"}]
    empty_mock = lambda text: []
    conclusion_mock = lambda chunks: [[], []]
    g, c, u, l = full_review_concurrently(
        "offline", samples, general_review=general_mock,
        conclusion_review=conclusion_mock,
        unrelated_review=empty_mock, language_review=empty_mock)
    assert len(g) == len(c) == len(u) == len(l) == 2
    assert [v[0]["quote"] for v in g] == ["Findings", "Conclusi"]
    assert c == u == l == [[], []]
    try:
        full_review_concurrently("offline", samples,
            general_review=lambda part: (_ for _ in ()).throw(
                ReviewFailure("GROQ_NETWORK")),
            conclusion_review=conclusion_mock,
            unrelated_review=empty_mock, language_review=empty_mock)
        raise AssertionError("Failed mandatory pass was ignored")
    except ReviewFailure as exc:
        assert exc.code == "GROQ_NETWORK"
    assert MODEL == "qwen/qwen3.8-27b"
    class FakeError:
        def read(self, limit):
            return b'{"error":{"message":"unsupported reasoning effort"}}'
    assert classify_http_400(FakeError()) == "GROQ_REQUEST_400_REASONING"
    class JsonFailed:
        def read(self, limit):
            return b'{"error":{"type":"invalid_request_error","code":"json_validate_failed","message":"generation failed"}}'
    assert classify_http_400(JsonFailed()) == "GROQ_REQUEST_400_JSON_GENERATION"
    class BadParams:
        def read(self, limit):
            return b'{"error":{"type":"invalid_request_error","message":"unknown parameter name"}}'
    assert classify_http_400(BadParams()) == "GROQ_REQUEST_400_PARAMETER"
    assert "Haronex-Radiology-Agent/11" in str(api_review.__code__.co_consts)
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
    def json_failed_once(request):
        called.append(1)
        if len(called) == 1:
            raise ReviewFailure("GROQ_REQUEST_400_JSON_GENERATION", True)
        return fake_response
    json_delays=[]
    assert api_review("gsk_dummy", sample, json_failed_once, json_delays.append) == []
    assert len(called) == 2 and json_delays == [1]
    called.clear()
    def always_json_failed(request):
        called.append(1)
        raise ReviewFailure("GROQ_REQUEST_400_JSON_GENERATION", True)
    try:
        api_review("gsk_dummy", sample, always_json_failed, json_delays.append)
        raise AssertionError("Excessive JSON generation retries")
    except ReviewFailure as exc:
        assert exc.code == "GROQ_REQUEST_400_JSON_GENERATION" and len(called) == 2
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

def unicode_protocol_self_test():
    """Windows cp1252/ANSI stdout must never destroy Groq's Unicode findings.
    The wire is ASCII-only JSON; JSON decoding restores the original Unicode.
    This is safe even when PyInstaller redirects console output to an ANSI pipe.
    """
    phrase = "مراجعة — Résumé – Left phleboli → phleboliths"
    example = {"ok": True, "mode": "groq", "passes": 2,
               "groups": [[{"quote": "phleboli", "category": "language",
                            "explanation": phrase,
                            "suggestion": "phleboliths"}]]}
    wire = json.dumps(example, ensure_ascii=True, separators=(",", ":"))
    assert wire.isascii()
    assert json.loads(wire)["groups"][0][0]["explanation"] == phrase
    print(wire, flush=True)


def serve():
    # Private stdin/stdout protocol; no listening ports, sockets or HTTP API.
    # Each request returns exactly one JSON line and flushes immediately.
    for raw_line in sys.stdin.buffer:
        if len(raw_line) > 900000:
            print(json.dumps({"ok": False, "error": "WORKER_REQUEST_TOO_LARGE"}), flush=True)
            continue
        try:
            line = raw_line.decode("utf-8", errors="strict")
        except UnicodeDecodeError:
            print(json.dumps({"ok": False, "error": "WORKER_INVALID_UTF8"}), flush=True)
            continue
        if not line.strip():
            continue
        main(line)
        sys.stdout.flush()


if __name__ == "__main__":
    if "--self-test" in sys.argv:
        self_test()
    elif "--unicode-self-test" in sys.argv:
        unicode_protocol_self_test()
    elif "--server" in sys.argv:
        serve()
    else:
        main()
