"""Test 3: edit a real hand-made team deck; untouched slides stay identical."""
import shutil
import time
from pathlib import Path

import pytest

import edit
from conftest import CASES, HERE, OUT


def test_edit_real_deck():
    src = CASES / "IE.pptx"
    if not src.exists():
        pytest.skip("case comp cases/IE.pptx not found")
    shutil.copyfile(src, OUT / "IE_copy.pptx")
    t = time.time()
    out, r = edit.edit(HERE / "fixtures" / "edit" / "edit_spec.yaml")
    assert time.time() - t < 60
    assert r["unchanged_ok"], r["changed_ids"]
    assert r["slides"] == 18 and r["added"] == 2


def test_add_one_slide_to_built_deck(golden_deck, tmp_path):
    import yaml
    deck = tmp_path / "deck.pptx"
    shutil.copyfile(golden_deck[0], deck)
    spec = tmp_path / "e.yaml"
    spec.write_text(yaml.safe_dump({"deck": "deck.pptx", "operations": [{"add": {
        "pattern": "text_content", "after": 30, "section": "Impact",
        "slots": {"title": "Funding comes from reinvested savings and one partner", "body": ["Savings of ~5m a year"],
                  "source": "Source: Team Analysis"}}}]}))
    t = time.time()
    out, r = edit.edit(spec)
    assert time.time() - t < 60
    assert r["unchanged_ok"] and r["slides"] == 39   # 37 + navigator + 1 new
    assert r.get("navigation"), "edit must refresh navigation on a navigated deck"


def test_edit_refuses_overwrite(tmp_path, golden_deck):
    import yaml
    deck = tmp_path / "deck.pptx"
    shutil.copyfile(golden_deck[0], deck)
    spec = tmp_path / "e.yaml"
    spec.write_text(yaml.safe_dump({"deck": "deck.pptx", "output": "deck.pptx", "operations": [{"delete": {"slide": 3}}]}))
    with pytest.raises(edit.EditError):
        edit.edit(spec)
