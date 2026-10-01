"""Optional local PDF receipt extraction, replacing the old Drive-specific script."""
from io import BytesIO
import re
import pdfplumber
from generator import InputError


def extract_receipt(data):
    if not data.startswith(b'%PDF') or len(data) > 10 * 1024 * 1024:
        raise InputError('Carga un PDF válido de hasta 10 MB.')
    try:
        with pdfplumber.open(BytesIO(data)) as pdf:
            if len(pdf.pages) > 20:
                raise InputError('El recibo puede tener como máximo 20 páginas.')
            text = '\n'.join(page.extract_text() or '' for page in pdf.pages)
    except InputError:
        raise
    except Exception as exc:
        raise InputError('No se pudo leer el PDF. Usa un PDF con texto y sin protección.') from exc
    if not text.strip():
        raise InputError('El PDF no contiene texto extraíble. Los escaneos requieren OCR.')
    patterns = {
        'SERIE_RECIBO': r'(?:N[°º]|Nro\.?\s*:?|N[uú]mero\s*:?)[ \n]*(E\d+\s*-\s*\d+)',
        'FECHA_EMISION': r'Fecha de [Ee]misi[oó]n(?:\s*Tipo de Moneda)?[\s:]*(\d{2}/\d{2}/\d{4})',
        'TOTAL': r'Total Neto Recibido\s*:?\s*(?:S/\.?\s*)?([\d,]+\.\d{2})',
    }
    fields = {key: (m.group(1).strip() if (m := re.search(pattern, text, re.IGNORECASE)) else '') for key, pattern in patterns.items()}
    return {'fields': fields, 'missing': [k for k, v in fields.items() if not v],
            'note': 'Extracción de texto para revisión manual; no valida el recibo.'}
