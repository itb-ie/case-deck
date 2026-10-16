#!/usr/bin/env python3
"""Quick install check. Run once after cloning:  uv sync && uv run scripts/selftest.py"""
from __future__ import annotations

import importlib
import platform
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

HERE = Path(__file__).resolve().parent
sys.path.insert(0, str(HERE))

OK, FAIL, WARN = "ok  ", "FAIL", "warn"


def line(status, msg):
    print(f"  [{status}] {msg}")
    return status != FAIL


def main():
    t0 = time.time()
    good = True
    print("case-deck self-test")
    good &= line(OK if sys.version_info >= (3, 9) else FAIL, f"Python {platform.python_version()} (3.9+ needed)")
    in_venv = sys.prefix != getattr(sys, "base_prefix", sys.prefix)
    line(OK if in_venv else WARN, "running in the skill's own environment" if in_venv
         else "running in the system Python: use `uv sync` and `uv run scripts/selftest.py`")
    for mod, pip in (("pptx", "python-pptx"), ("lxml", "lxml"), ("PIL", "Pillow"), ("yaml", "PyYAML")):
        try:
            importlib.import_module(mod)
            good &= line(OK, f"{pip} installed")
        except ImportError:
            good &= line(FAIL, f"{pip} missing: run `uv sync` in the skill folder")
    try:
        importlib.import_module("pypdf")
        line(OK, "pypdf installed (PowerPoint rendering)")
    except ImportError:
        line(WARN, "pypdf missing: only needed for --engine powerpoint")

    import deckkit as dk
    for f in (dk.TEMPLATE, dk.PATTERNS, dk.CATALOG):
        good &= line(OK if f.exists() else FAIL, f"{f.relative_to(dk.SKILL_DIR)}")

    try:
        import ai_check
        cmd = ai_check.detector_cmd()
        local = "node_modules" in cmd[0]
        line(OK if local else WARN, "AI detector " + ("installed locally" if local else "via npx (run `npm install` to work offline)"))
    except Exception as e:
        good &= line(FAIL, f"AI detector: {e}")

    renderer = "Quick Look" if shutil.which("qlmanage") else ("PowerPoint" if platform.system() == "Windows" else None)
    line(OK if renderer else WARN, f"renderer: {renderer or 'none (renders and thumbnails are skipped)'}")

    import build
    import review
    with tempfile.TemporaryDirectory() as td:
        spec = Path(td) / "deck_spec.yaml"
        spec.write_text("""
deck: {title: Self-test, client: Testco, output: selftest.pptx}
slides:
- pattern: exec_summary
  section: Executive Summary
  slots:
    question: How can Testco double its export revenue by 2030?
    c1_head: Win abroad
    c2_head: Sell online
    c1_question: Which three markets should Testco enter first?
    c2_question: How does Testco reach buyers online?
    c1_reco: ["**Three-market pilot**", "Enter Spain, Mexico and Chile with local partners"]
    c2_reco: ["**Direct online shop**", "Sell the top 20 products online with 48-hour delivery"]
    impact: "Impact: USD ~12m extra revenue by 2030"
    source: "Source: Team Analysis"
- pattern: text_content
  section: Analysis
  slots:
    title: Testco grows 8% a year at home but exports only 5% of sales
    body: ["Home market growth of 8% a year since 2020", "Exports are 5% of sales, against 30% for peers"]
    source: "Source: Team Analysis"
- pattern: closing
  slots: {m1_name: Ana, m2_name: Ben}
""", encoding="utf-8")
        try:
            out, rep = build.build(spec, quiet=True)
            good &= line(OK, f"built a {rep['slides']}-slide sample in {rep['seconds']}s")
            r = review.review(out, run_ai=False)
            highs = [f for f in r["findings"] if f["severity"] == "high" and f["category"] != "structure"]
            good &= line(OK if not highs else FAIL, f"review ran ({len(r['findings'])} findings, {len(highs)} high)")
        except Exception as e:
            good &= line(FAIL, f"build or review failed: {e}")
    print(f"\n{'All good' if good else 'Problems found (see FAIL lines)'} in {time.time() - t0:.1f}s")
    sys.exit(0 if good else 1)


if __name__ == "__main__":
    main()
