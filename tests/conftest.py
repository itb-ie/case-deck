"""Shared test setup. The winning decks and the case files live outside the
skill repo; tests that need them are skipped when they are not found.
Set CASE_DECK_WINNING and CASE_DECK_CASES to point elsewhere."""
from __future__ import annotations

import os
import sys
from pathlib import Path

import pytest

HERE = Path(__file__).resolve().parent
ROOT = HERE.parent
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(HERE))

WINNING = Path(os.environ.get("CASE_DECK_WINNING", ROOT.parent / "winning_cases"))
CASES = Path(os.environ.get("CASE_DECK_CASES", ROOT.parent / "case comp cases"))
OUT = HERE / "output"
OUT.mkdir(exist_ok=True)


def _mtimes():
    files = []
    for d in (WINNING, ROOT.parent / "prompts" / "reference", ROOT / "assets"):
        if d.exists():
            files += [p for p in d.iterdir() if p.is_file()]
    return {str(p): p.stat().st_mtime for p in files}


@pytest.fixture(scope="session", autouse=True)
def originals_untouched():
    """Test 8: nothing in winning_cases/, prompts/reference/ or assets/ changes."""
    before = _mtimes()
    yield
    after = _mtimes()
    changed = [p for p in before if after.get(p) != before[p]]
    assert not changed, f"source files were modified: {changed}"


@pytest.fixture(scope="session")
def winning():
    if not (WINNING / "Mexico 24H Case.pptx").exists():
        pytest.skip("winning decks not found (set CASE_DECK_WINNING)")
    return WINNING


@pytest.fixture(scope="session")
def golden_deck(winning):
    import build
    import make_golden
    spec_path = HERE / "fixtures" / "tajin" / "deck_spec.yaml"
    values = make_golden.extract(winning)       # images into fixtures/tajin/img (git-ignored)
    if not spec_path.exists():
        import yaml
        spec_path.write_text(yaml.safe_dump(make_golden.make_spec(values), allow_unicode=True, sort_keys=False))
    out, report = build.build(spec_path, OUT / "tajin_rebuild.pptx", quiet=True)
    return out, report
