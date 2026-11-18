"""Offline performance and clinical-quality contract. No API key, PHI or network."""
import importlib.util
import io
import json
import os
import subprocess
import sys
import time

MODULE = os.path.join(os.path.dirname(__file__), "GroqReviewWorker.py")
spec = importlib.util.spec_from_file_location("haronex_worker", MODULE)
w = importlib.util.module_from_spec(spec)
spec.loader.exec_module(w)

def run():
    texts = [
        "Findings: Left renal calculus. Mild hepatomegaly. phleboli.\r",
        "Conclusion: No renal calculus. MRI MRCP IHBR 2026/2027.\r"
    ]
    calls = {"general": 0, "language": 0, "unrelated": 0, "clinical": 0}
    def review(piece):
        calls["general"] += 1
        time.sleep(.04)
        if "phleboli" in piece:
            return [{"quote": "phleboli", "category": "language",
                     "explanation": "Spelling", "suggestion": "phleboliths"}]
        return []
    def language(piece):
        calls["language"] += 1
        time.sleep(.04)
        return []
    def unrelated(piece):
        calls["unrelated"] += 1
        time.sleep(.04)
        return []
    def clinical(chunks):
        calls["clinical"] += 1
        time.sleep(.04)
        return [[{"quote": "No renal calculus", "category": "logic",
                  "explanation": "Contradicts calculus in Findings",
                  "suggestion": "Verify with radiologist",
                  "section": "conclusion"}] if "No renal calculus" in p else []
                for p in chunks]
    t0=time.perf_counter()
    first=w.full_review_concurrently("gsk_offline_testing_1234",texts,
        general_review=review,conclusion_review=clinical,
        unrelated_review=unrelated,language_review=language)
    elapsed=time.perf_counter()-t0
    assert len(first)==4 and all(len(item)==2 for item in first)
    assert first[0][0][0]["quote"]=="phleboli"
    assert first[1][1][0]["category"]=="logic"
    assert calls=={"general":2,"language":2,"unrelated":2,"clinical":1}
    assert elapsed < .30, ("Parallel audit unexpectedly serialized",elapsed)

    # Cache only SAME exact text+pass; never cache Conclusion cross-section.
    w.reset_metrics()
    key="gsk_offline_testing_1234"
    captured = []
    for part in texts:
        captured.append(w.SECTION_CACHE.get_or_run(
            "general_test", key,part,lambda text: review(text)))
    before=calls["general"]
    for part in texts:
        reused=w.SECTION_CACHE.get_or_run(
            "general_test",key,part,lambda text: review(text))
    assert calls["general"]==before and w.snapshot_metrics()["cache_hits"]>=2
    changed=texts[0].replace("phleboli","phleboliths")
    w.SECTION_CACHE.get_or_run("general_test",key,changed,lambda text: review(text))
    assert calls["general"]==before+1
    clinical(texts)
    clinical([changed,texts[1]])
    assert calls["clinical"]==3, "Clinical context pass must never be skipped"

    # All mandatory passes fail closed rather than returning partial findings.
    try:
        w.full_review_concurrently(key,texts,general_review=review,
            conclusion_review=lambda x: (_ for _ in ()).throw(
                w.ReviewFailure("GROQ_NETWORK")),
            unrelated_review=unrelated,language_review=language)
        raise AssertionError("Failed clinical pass was accepted")
    except w.ReviewFailure as exc:
        assert exc.code=="GROQ_NETWORK"
    # Adaptive rate limiting and leak-free condition release.
    limiter=w.AdaptiveLimiter()
    assert limiter.limit==4
    limiter.acquire()
    limiter.release(rate_limited=True)
    assert limiter.limit==2 and limiter.active==0
    for _ in range(10):
        limiter.acquire()
        limiter.release(success=True)
    assert limiter.limit==3 and limiter.active==0

    # V12 zero-Logic regression: a deterministic positive-vs-negative
    # contradiction survives independently of stochastic Qwen outputs.
    contradictory = ("Findings:\rBilateral renal cortical cysts detected.\r"
                     "Conclusion:\rNo renal cysts.\r")
    matched=w.independent_direct_contradiction_groups([contradictory])
    assert len(matched[0])==1 and matched[0][0]["category"]=="logic"
    assert not w.independent_direct_contradiction_groups([
        "Findings:\rNo renal cysts.\rConclusion:\rNo renal cysts.\r"])[0]
    # Explicit patient metadata warning, no gender guessing from names.
    sex_mismatch=w.patient_anatomy_consistency_groups([
        "Sex: Male\rFindings:\rUterine mass is present.\r"
        "Conclusion:\rUterine mass.\r"])
    assert len(sex_mismatch[0])==1
    assert not w.patient_anatomy_consistency_groups([
        "Patient: Alex\rFindings:\rUterine mass.\rConclusion:\rMass.\r"])[0]
    # Clinical omissions and certainty mismatch must produce real Logic
    # cards, each tied to an exact quote, even with AI disabled.
    pet = ("Patient data: Male patient 78 years old.\\r"
           "Clinical background: History of ? prostatic cancer.\\r"
           "Findings:\\rProgression of the mild pericardial effusion.\\r"
           "Conclusion:\\rHistory of prostatic malignancy on hormonal therapy.\\r")
    coverage = w.independent_clinical_coverage_groups([pet])[0]
    assert len(coverage) == 2, coverage
    assert all(row["category"] == "logic" and row["quote"] in pet
               for row in coverage)
    assert any(row["section"] == "conclusion" for row in coverage)
    assert any(row["section"] == "body" for row in coverage)
    # A persistent worker serves TWO separate requests in one process; input
    # invalid on purpose so the smoke never touches network.
    binary=sys.argv[1] if len(sys.argv)>1 else None
    command=([binary,"--server"] if binary else
             [sys.executable,MODULE,"--server"])
    proc=subprocess.run(command,
        input='{}\n{}\n',text=True,capture_output=True,timeout=20,
        encoding="utf-8")
    replies=[json.loads(line) for line in proc.stdout.splitlines() if line.strip()]
    assert proc.returncode==0 and len(replies)==2
    assert all(x.get("ok") is False for x in replies)
    assert all("texts" not in x and "api_key" not in x for x in replies)
    print(json.dumps({"ok":True,"test":"v12_evidence_first_quality_offline",
                      "four_passes":True,"persistent_protocol":True,
                      "section_cache":True,"clinical_rerun":True,
                      "adaptive_scheduler":True}))
if __name__=="__main__":
    run()
