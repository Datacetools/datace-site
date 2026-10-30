import hashlib
import json
import re
import urllib.parse
from f24_core import export_excel_bytes, parse_pdf, periodo_versamenti, totali
MAX_FILES = 10
MAX_SIZE_MB = 20
MAX_SIZE_BYTES = MAX_SIZE_MB * 1024 * 1024
MAX_TOTAL_MB = 100
MAX_TOTAL_BYTES = MAX_TOTAL_MB * 1024 * 1024
VERSIONE = '1.03'
XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

class Risposta(Exception):

    def __init__(self, status, detail):
        super().__init__(detail)
        self.status, self.detail = (status, detail)

def _errore(status, error_code, stage, message):
    return Risposta(status, {'error_code': error_code, 'stage': stage, 'message': message})

def _page_bucket(n):
    if not n or n <= 0:
        return 'unknown'
    if n == 1:
        return '1'
    if n <= 5:
        return '2-5'
    return '6+'

def _text_layer_flag(text_pages, total_pages):
    if not total_pages:
        return 'unknown'
    if text_pages == 0:
        return 'false'
    if text_pages >= total_pages:
        return 'true'
    return 'partial'

def _elabora(nomi, tipi, files, credito_negativo):
    if sum((len(f) for f in files)) > MAX_TOTAL_BYTES:
        raise Risposta(413, {'error_code': 'file_too_large', 'stage': 'upload_validation', 'message': f'Troppi dati in un invio solo (massimo {MAX_TOTAL_MB} MB): dividi i file in più invii.'})
    if not files:
        raise _errore(400, 'empty_or_corrupt', 'upload_validation', 'Nessun file ricevuto.')
    if len(files) > MAX_FILES:
        raise _errore(400, 'too_many_files', 'upload_validation', f'Massimo {MAX_FILES} file per volta.')
    all_rows, parse_errors, total_pages, text_pages, pdf_letti = ([], 0, 0, 0, 0)
    impronte, doppioni, controllate, non_tornano = (set(), 0, 0, [])
    for posizione, (nome, tipo, raw) in enumerate(zip(nomi, tipi, files)):
        raw = bytes(raw)
        if tipo not in ('application/pdf', 'application/octet-stream'):
            if not (nome or '').lower().endswith('.pdf'):
                raise _errore(400, 'not_a_pdf', 'upload_validation', 'Uno dei file scelti non è un PDF valido.')
        if len(raw) > MAX_SIZE_BYTES:
            raise _errore(400, 'file_too_large', 'upload_validation', f'Uno dei file supera i {MAX_SIZE_MB} MB consentiti.')
        if len(raw) < 4:
            raise _errore(400, 'empty_or_corrupt', 'upload_validation', 'Uno dei file è vuoto o corrotto.')
        if not raw.startswith(b'%PDF'):
            raise _errore(400, 'not_a_pdf', 'upload_validation', 'Uno dei file scelti non è un PDF valido.')
        impronta = hashlib.sha256(raw).digest()
        doppioni += impronta in impronte
        impronte.add(impronta)
        try:
            stats = {}
            rows = parse_pdf(raw, stats=stats)
            total_pages += stats.get('pages', 0)
            text_pages += stats.get('text_pages', 0)
            pdf_letti += bool(rows)
            controllate += stats.get('deleghe_controllate', 0)
            pagine = stats.get('pagine_non_tornano', [])
            if pagine:
                non_tornano.append({'file': posizione, 'pagine': pagine[:50], 'n': len(pagine)})
            all_rows.extend(rows)
        except Exception:
            parse_errors += 1
    text_flag = _text_layer_flag(text_pages, total_pages)
    if not all_rows:
        if text_flag == 'false':
            raise _errore(422, 'no_text_layer', 'data_parse', 'Il PDF non contiene testo leggibile (probabile scansione).')
        if parse_errors:
            raise _errore(422, 'unsupported_layout', 'data_parse', 'Impossibile leggere i file: assicurati che siano F24 PDF non cifrati.')
        raise _errore(422, 'parse_no_data', 'data_parse', 'Non sono stati trovati dati F24 nel documento.')
    try:
        excel = export_excel_bytes(all_rows, credito_negativo=credito_negativo)
    except Exception:
        raise _errore(500, 'server_error', 'excel_gen', 'Errore nella generazione del file Excel.')
    if len(files) > 1:
        fname = 'f24_export.xlsx'
    else:
        raw_name = (nomi[0] or 'f24').replace('.pdf', '')
        safe_name = re.sub('[^a-zA-Z0-9._\\-]', '_', raw_name)[:50]
        fname = f'{safe_name}_f24.xlsx'
    dal, al = periodo_versamenti(all_rows)
    riepilogo = {**totali(all_rows), 'pdf_letti': pdf_letti, 'doppioni': doppioni, 'dal': dal, 'al': al, 'controllate': controllate, 'non_tornano': non_tornano}
    return (excel, {'Content-Disposition': f"attachment; filename*=UTF-8''{urllib.parse.quote(fname, safe='._-')}", 'X-Rows-Extracted': str(len(all_rows)), 'X-Files-Processed': str(len(files) - parse_errors), 'X-Files-Failed': str(parse_errors), 'X-Page-Count-Bucket': _page_bucket(total_pages), 'X-Has-Text-Layer': text_flag, 'X-Tool-Version': VERSIONE, 'X-Riepilogo': json.dumps(riepilogo)})

def elabora(nomi, tipi, files, credito_negativo):
    try:
        excel, headers = _elabora(list(nomi), list(tipi), list(files), bool(credito_negativo))
    except Risposta as e:
        return (json.dumps({'status': e.status, 'detail': e.detail}), b'')
    return (json.dumps({'status': 200, 'headers': headers, 'tipo': XLSX}), excel)
