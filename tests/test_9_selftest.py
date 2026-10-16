"""Test 9: the student self-test passes in under 30 seconds."""
import subprocess
import sys
import time

from conftest import ROOT


def test_selftest():
    t = time.time()
    r = subprocess.run([sys.executable, str(ROOT / "scripts" / "selftest.py")], capture_output=True, text=True, timeout=60)
    assert r.returncode == 0, r.stdout + r.stderr
    assert time.time() - t < 30
