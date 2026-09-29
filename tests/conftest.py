from __future__ import annotations

import sys
from pathlib import Path

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT / "scripts"))
sys.path.insert(0, str(ROOT / "tests"))

import pytest

from create_test_documents import create_test_documents


@pytest.fixture
def documents(tmp_path):
    return create_test_documents(tmp_path)
