import pytest
from receipts import extract_receipt
from generator import InputError
from pathlib import Path


def test_receipt_patterns_and_missing(monkeypatch):
    class Page:
        def extract_text(self):
            return 'N° E001-123\nFecha de Emisión 01/10/2026\nTotal Neto Recibido: S/ 2,400.00'
    class PDF:
        pages = [Page()]
        def __enter__(self): return self
        def __exit__(self, *args): pass
    monkeypatch.setattr('receipts.pdfplumber.open', lambda *_: PDF())
    result = extract_receipt(b'%PDF sample')
    assert result['fields'] == {'SERIE_RECIBO':'E001-123','FECHA_EMISION':'01/10/2026','TOTAL':'2,400.00'}
    assert result['missing'] == []
    monkeypatch.setattr(Page, 'extract_text', lambda _: 'Texto sin campos')
    assert len(extract_receipt(b'%PDF sample')['missing']) == 3


def test_fake_pdf_rejected():
    with pytest.raises(InputError): extract_receipt(b'not pdf')


def test_actual_pdf_fixture():
    result = extract_receipt((Path(__file__).parent / 'fixtures/receipt.pdf').read_bytes())
    assert result['missing'] == []
    assert result['fields']['TOTAL'] == '2,400.00'
