"""Optional real-document regressions; keep uploaded scans outside the checkout.

SCAN_TEST_SAMPLE_DIR=/path/to/private/scans python -m pytest -q tests/test_supplied_scans.py
Set SCAN_TEST_ITEM_LIST to validate using an actual workbook. Without it, the
temporary catalog below is transcribed from the document headers.
"""

from datetime import datetime
import hashlib
import os
from pathlib import Path
import shutil

import pytest
from pypdf import PdfReader

from test_scan_sorter import sorter


SAMPLE_DIR = os.environ.get("SCAN_TEST_SAMPLE_DIR")
ITEM_LIST = os.environ.get("SCAN_TEST_ITEM_LIST")
pytestmark = pytest.mark.skipif(not SAMPLE_DIR, reason="Private scan samples not supplied")
SAMPLES = [
    ("20261007-13100747-UNKNOWN-NOPO_1.pdf", "13100747", "CINNAMON OR EXTRACT WS NAT", "PO005157", 8),
    ("20261007-NO_CODE_FOUND-UNKNOWN-STOCK_1.pdf", "13100029", "ESPRESSO EXTRACT WS", "PO005191", 8),
    ("20261007-NO_CODE_FOUND-UNKNOWN-PO005157_1.pdf", "13100748", "CINNAMON, TOASTED WONF WS NAT", "PO005157", 7),
    ("20261007-43100228-UNKNOWN-NOPO_1.pdf", "13100228", "CHOCOLATE TYPE RXN WS NAT (2X)", "PO031440", 8),
]


def sample_item_map():
    if ITEM_LIST:
        return sorter.load_item_map(Path(ITEM_LIST))
    return {entry[1]: entry[2] for entry in SAMPLES}


@pytest.mark.parametrize("filename, code, description, po, page_count", SAMPLES)
def test_supplied_scan_fields_and_pdf_preservation(tmp_path, monkeypatch,
                                                  filename, code, description, po, page_count):
    monkeypatch.delenv("SCAN_SMTP_PASSWORD", raising=False)
    notices = []
    monkeypatch.setattr(sorter, "send_error_email", lambda *args: notices.append(args))
    incoming, output, debug = [tmp_path / p for p in ("input", "output", "debug")]
    for directory in (incoming, output, debug):
        directory.mkdir()
    sample = Path(SAMPLE_DIR) / filename
    original_hash = hashlib.sha256(sample.read_bytes()).digest()
    source = incoming / filename
    shutil.copyfile(sample, source)
    config = sorter.Config(incoming, output, debug,
                           Path(ITEM_LIST) if ITEM_LIST else tmp_path / "test-catalog.xlsx")
    item_map = sample_item_map()
    result = sorter.process_file(source, config, item_map, datetime(2026, 10, 7))
    assert result.destination.name == f"20261007-{code}-{description}-{po}_1.pdf"
    assert not result.reasons
    assert not notices
    assert len(PdfReader(result.destination).pages) == page_count
    assert hashlib.sha256(result.destination.read_bytes()).digest() == original_hash
    assert hashlib.sha256(sample.read_bytes()).digest() == original_hash
    assert not source.exists()


def test_upside_down_checklist_also_uses_field_cells(tmp_path):
    sample = Path(SAMPLE_DIR) / SAMPLES[0][0]
    config = sorter.Config(tmp_path, tmp_path, tmp_path, tmp_path / "test-catalog.xlsx")
    texts = {}
    items = sample_item_map()
    with sorter.read_first_page(sample, config) as image:
        with image.rotate(180, expand=True, fillcolor="white") as upside_down:
            fields = sorter.recognize_image(upside_down, items, config, texts, tmp_path)
    assert fields == ("13100747", "PO005157")
    assert any(key.startswith("po-cell-rotated-") for key in texts)
