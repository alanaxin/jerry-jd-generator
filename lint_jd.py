#!/usr/bin/env python3
"""
lint_jd.py — deterministic linter for Jerry JDs.
Reads a .docx, runs zero-/low-judgment checks against the Hard Rules, emits JSON.

Covers (deterministic): #1 X-not-Y, #9 from-day-one, #10 corporate/AI filler,
#13 nested bullets, #17 quoted archetypes, #19 year ranges, #20 comp/benefits.
Pattern-flags for auditor follow-up: #2 cadence repetition, #8 not-sexy opener,
#11 redundant 'its own' triples, #15 vague quantifiers, #16 category mix,
#18 possible personal names, #22 junior/high-supply heading.

Each finding: {rule, severity, line, text, note}
severity: "fail" = deterministic violation; "flag" = needs auditor judgment.

Rules NOT covered (model/human judgment only):
  #3 setup-line throat-clear, #4 punchy fragments, #5 X-answered/Y-open,
  #6 conclusory sentences, #7 fake parallelism, #12 list ordering,
  #14 subheading fit, #21 differentiators vs table stakes,
  #23 factual fidelity, #24 confidentiality ceiling,
  #25 cross-bullet restatement, #26 Radical Honesty friction test
"""
import sys, re, json, subprocess, os

def extract_md(path):
    """Extract text from a docx. Tries extract-text, then pandoc, then python-docx."""
    # Try extract-text (available in Claude Code environments)
    out = subprocess.run(["extract-text", path], capture_output=True, text=True)
    if out.returncode == 0 and out.stdout.strip():
        return out.stdout
    # Try pandoc (available if installed: brew install pandoc)
    out = subprocess.run(["pandoc", path, "-t", "markdown"], capture_output=True, text=True)
    if out.returncode == 0 and out.stdout.strip():
        return out.stdout
    # Fallback: python-docx (pip3 install python-docx)
    try:
        from docx import Document
        doc = Document(path)
        lines = []
        for para in doc.paragraphs:
            if para.style.name.startswith("Heading"):
                level = para.style.name.split(" ")[-1]
                try:
                    prefix = "#" * int(level)
                except ValueError:
                    prefix = "#"
                lines.append(f"{prefix} {para.text}")
            elif para.text.strip():
                if "List" in para.style.name:
                    # Detect list indent level via numPr ilvl so Rule 13 works correctly.
                    # ilvl=0 = top-level bullet, ilvl>=1 = nested sub-bullet.
                    indent = 0
                    try:
                        pPr = para._p.pPr
                        if pPr is not None:
                            numPr = pPr.numPr
                            if numPr is not None and numPr.ilvl is not None:
                                indent = int(numPr.ilvl.val)
                    except Exception:
                        pass
                    # Prefix with 2 spaces per indent level so indent_level() fires correctly.
                    prefix_spaces = "  " * indent
                    parts = []
                    for run in para.runs:
                        if run.bold:
                            parts.append(f"**{run.text}**")
                        else:
                            parts.append(run.text)
                    lines.append(f"{prefix_spaces}- " + "".join(parts))
                else:
                    lines.append(para.text)
        return "\n".join(lines)
    except ImportError:
        raise RuntimeError(
            "No text extraction method available.\n"
            "Run: pip3 install python-docx --break-system-packages"
        )

def clean(s):
    return s.replace("“", '"').replace("”", '"').replace("’", "'").replace("‘", "'")

def get_lines(md):
    return [(i+1, ln) for i, ln in enumerate(md.splitlines()) if ln.strip()]

def is_heading(ln): return ln.lstrip().startswith("#")
def is_bullet(ln):  return bool(re.match(r'^\s*[-*]\s', ln))
def bold_label(ln):
    m = re.match(r'^\s*[-*]\s+\*\*(.+?)\*\*', ln)
    return m.group(1).strip() if m else None
def indent_level(ln):
    # Returns the number of leading spaces before the bullet marker.
    # 0 = top-level; >=2 = nested (ilvl>=1 from docx produces 2 spaces per level).
    m = re.match(r'^(\s*)[-*]\s', ln)
    return len(m.group(1)) if m else 0

# ---- Rule 10: corporate filler + AI-default prose patterns ----
# Corporate filler: generic HR phrases
CORPORATE_FILLER = [
    "dynamic environment",
    "wear many hats",
    "passionate about innovation",
    "fast-paced environment",
    "fast-paced and dynamic",
    "think outside the box",
    "self-starter",
    "hit the ground running",
    "best-in-class",
    "synergy",
    "rockstar",
    "ninja",
]
# AI-default patterns: phrases statistically common in AI-generated text
AI_DEFAULT_PATTERNS = [
    "ongoing journey",
    "evolving landscape",
    "ever-changing",
    "at the intersection of",
    "a unique opportunity to",
    "in a rapidly evolving",
    "fast-paced and ever",
]

STALE_NOTSEXY_CATS = [
    "payroll", "accounting", "compliance", "bookkeeping",
    "data entry", "administrative", "back office", "back-office",
]

def lint(path):
    md = extract_md(path)
    findings = []
    lines = get_lines(md)

    current_section = None

    for idx, raw in lines:
        ln = clean(raw)
        low = ln.lower()

        if is_heading(ln):
            current_section = low.lstrip("# ").strip()
            # Rule 22: "why we need you" heading flag
            if "why we need you" in current_section:
                findings.append({"rule": 22, "severity": "flag", "line": idx,
                    "text": ln.strip(),
                    "note": "'Why we need you' heading — banned for junior/high-supply roles; auditor must confirm role seniority warrants it."})
            continue

        # ---- Rule 1: X-not-Y / antithesis ----
        for m in re.finditer(r'(\b[\w\-]+),\s+not\s+([\w\-]+)', ln):
            findings.append({"rule": 1, "severity": "fail", "line": idx,
                "text": ln.strip(),
                "note": f"X-not-Y construction: '{m.group(0)}'. Banned in all forms."})
        if re.search(r'\bstops?\s+\w+ing\s+and\s+starts?\s+\w+ing\b', low):
            findings.append({"rule": 1, "severity": "fail", "line": idx,
                "text": ln.strip(),
                "note": "Antithesis cadence ('stops X-ing and starts Y-ing')."})

        # ---- Rule 17: quoted archetype labels ----
        lbl = bold_label(ln)
        if lbl is not None:
            if re.match(r'^["\']', lbl) or re.search(r'["\']\s*:?$', lbl):
                findings.append({"rule": 17, "severity": "fail", "line": idx,
                    "text": lbl,
                    "note": "Archetype label uses quotation marks. Strip them. No exceptions."})

        # ---- Rule 13: nested sub-bullets ----
        # indent_level returns leading-space count; >=2 means ilvl>=1 in docx (nested).
        if is_bullet(raw) and indent_level(raw) >= 2:
            findings.append({"rule": 13, "severity": "fail", "line": idx,
                "text": ln.strip()[:80],
                "note": "Nested sub-bullet detected. Flat lists only."})

        # ---- Rule 19: years range / ceiling ----
        if re.search(r'\b\d+\s*(?:to|[-–—])\s*\d+\s+years', low) or \
           re.search(r'\b(?:up to|maximum of|no more than)\s+\d+\s+years', low):
            findings.append({"rule": 19, "severity": "fail", "line": idx,
                "text": ln.strip(),
                "note": "Years-of-experience range/ceiling. Use minimum-only ('X+ years')."})

        # ---- Rule 20: comp / benefits ----
        if re.search(r'\$\s?\d[\d,]*\s*(?:k|,000|to|–|-|\bplus\b)?', ln) and \
           re.search(r'salary|compensation|equity|bonus|benefits|pays?\b|\bpay\b', low):
            findings.append({"rule": 20, "severity": "fail", "line": idx,
                "text": ln.strip(),
                "note": "Compensation/benefits language. Handled in ATS, never in a JD."})
        elif re.search(r'\b(competitive salary|salary range|equity package|401\(?k\)?|health benefits|stock options)\b', low):
            findings.append({"rule": 20, "severity": "fail", "line": idx,
                "text": ln.strip(),
                "note": "Compensation/benefits language. Handled in ATS, never in a JD."})

        # ---- Rule 18: names instead of titles ----
        PLACE_WORDS = {"North","South","Latin","Central","New","Silicon","United","Great","Bay","San","Los","Las","Hong","Costa"}
        for m in re.finditer(r'\b(?:report to|reports to|reporting to|led by|managed by|manager,)\s+([A-Z][a-z]+)\s+([A-Z][a-z]+)\b', ln):
            first = m.group(1)
            if first in PLACE_WORDS:
                continue
            cand = f"{m.group(1)} {m.group(2)}"
            findings.append({"rule": 18, "severity": "flag", "line": idx,
                "text": cand,
                "note": f"Possible personal name '{cand}' — JDs use job titles only. Confirm and replace with title."})

        # ---- Rule 2: cadence repetition ----
        m = re.search(r'\bwe\s+(\w+)\b.*\band\s+we\s+(\1)\b', low)
        if m:
            findings.append({"rule": 2, "severity": "flag", "line": idx,
                "text": ln.strip(),
                "note": f"Repeated verb '{m.group(1)}' in parallel clauses — cadence-over-content. Auditor confirm."})

        # ---- Rule 11: 'its own X, its own Y' triples ----
        if low.count("its own") >= 2:
            findings.append({"rule": 11, "severity": "flag", "line": idx,
                "text": ln.strip(),
                "note": "Repeated 'its own ...' — likely redundant list items making one point (#11). Collapse."})

        # ---- Rule 8: stale 'not sexy/glamorous' opener ----
        if re.search(r'\bnot\s+(?:sexy|glamorous|exciting)\b', low):
            cat_hit = next((c for c in STALE_NOTSEXY_CATS if c in low), None)
            sev = "fail" if cat_hit else "flag"
            note = ("Stale 'not sexy/glamorous' for a category that's already the default punchline"
                    + (f" ('{cat_hit}')" if cat_hit else "") + ". " if cat_hit
                    else "'not sexy/glamorous' opener — auditor must judge freshness for this category (#8).")
            findings.append({"rule": 8, "severity": sev, "line": idx,
                "text": ln.strip(), "note": note})

        # ---- Rule 15: vague quantifiers ----
        for m in re.finditer(r'\b(more than a dozen|a dozen|several|numerous|many|a handful of|dozens of)\b', low):
            findings.append({"rule": 15, "severity": "flag", "line": idx,
                "text": ln.strip(),
                "note": f"Vague quantifier '{m.group(1)}'. Use the real number if known (#15)."})

        # ---- Rule 16: category mix ('states plus Canada') ----
        if re.search(r'states?\s+plus\s+[A-Z][a-z]+', ln) or re.search(r'\bplus\s+canada\b', low):
            findings.append({"rule": 16, "severity": "flag", "line": idx,
                "text": ln.strip(),
                "note": "Category mix: sub-national units and a country in one list (#16)."})

        # ---- Rule 10: corporate filler ----
        for f in CORPORATE_FILLER:
            if f in low:
                findings.append({"rule": 10, "severity": "fail", "line": idx,
                    "text": ln.strip(),
                    "note": f"Corporate filler: '{f}' (#10)."})

        # ---- Rule 10: AI-default prose patterns ----
        for f in AI_DEFAULT_PATTERNS:
            if f in low:
                findings.append({"rule": 10, "severity": "fail", "line": idx,
                    "text": ln.strip(),
                    "note": f"AI-default pattern: '{f}' (#10). Replace with the specific thing."})

        # ---- Rule 9: 'from day one' ----
        if "from day one" in low:
            findings.append({"rule": 9, "severity": "fail", "line": idx,
                "text": ln.strip(),
                "note": "'from day one' filler (#9)."})

        # ---- Rule 27: bold label brevity (Sections 2 & 4) ----
        lbl27 = bold_label(ln)
        if lbl27 is not None:
            word_count = len(lbl27.split())
            if word_count > 4:
                findings.append({"rule": 27, "severity": "flag", "line": idx,
                    "text": lbl27,
                    "note": f"Bold label is {word_count} words — target ≤4. Cut every non-essential word (#27)."})

        # ---- Rule 28: no descriptive sub-heading within a phase ----
        # Flags patterns like "Reddit (Days 1-90)" or "Expanding Channel Presence"
        # appearing as a bold-only bullet or heading immediately after Phase 1/2 content.
        if is_heading(ln) and re.search(r'\bdays?\s+\d', low):
            findings.append({"rule": 28, "severity": "fail", "line": idx,
                "text": ln.strip(),
                "note": "Descriptive temporal sub-heading inside a phase (e.g. 'Reddit (Days 1–90)'). Delete it — the phase label is sufficient (#28)."})

    summary = {
        "fail": sum(1 for f in findings if f["severity"] == "fail"),
        "flag": sum(1 for f in findings if f["severity"] == "flag"),
        "total": len(findings),
    }
    return {"file": os.path.basename(path), "summary": summary, "findings": findings}

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("usage: lint_jd.py <file.docx>", file=sys.stderr); sys.exit(2)
    result = lint(sys.argv[1])
    print(json.dumps(result, indent=2, ensure_ascii=False))
