"""Test 7: every output reloads with no orphan parts; on macOS with
PowerPoint, PowerPoint opens and exports it (no repair dialog)."""
import platform
from pathlib import Path

import pytest

import deckkit as dk
from conftest import OUT

OUTPUTS = ["tajin_rebuild.pptx", "tajin_nav.pptx", "seeded.pptx", "ai_default_fixed.pptx", "IE_edited.pptx"]


@pytest.mark.parametrize("name", OUTPUTS)
def test_reloads(name):
    p = OUT / name
    if not p.exists():
        pytest.skip(f"{name} not built in this run")
    assert not dk.verify_package(p)


@pytest.mark.parametrize("name", ["tajin_nav.pptx", "IE_edited.pptx"])
def test_powerpoint_opens(name, tmp_path):
    p = OUT / name
    if platform.system() != "Darwin" or not Path("/Applications/Microsoft PowerPoint.app").exists():
        pytest.skip("PowerPoint for Mac not available")
    if not p.exists():
        pytest.skip(f"{name} not built in this run")
    import render
    pdf = render.powerpoint_pdf(p, tmp_path / "out.pdf", timeout=180)
    assert pdf.exists() and pdf.stat().st_size > 10_000
