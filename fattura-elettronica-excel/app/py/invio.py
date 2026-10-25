import io
import json
import re
import urllib.parse
import zipfile
from fattura_core import FormatoNonSupportato, estrai_file, export_excel_bytes, parse_fattura
MAX_FILES = 50
MAX_SIZE_MB = 10
MAX_SIZE_BYTES = MAX_SIZE_MB * 1024 * 1024
MAX_TOTAL_MB = 100
MAX_TOTAL_BYTES = MAX_TOTAL_MB * 1024 * 1024
XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

class Risposta(Exception):

    def __init__(self, status, detail):
        super().__init__(detail)
        self.status, self.detail = (status, detail)

def _is_valid_name(name):
    n = (name or '').lower()
    return n.endswith('.xml') or n.endswith('.p7m')
_letti = {'chiave': None}

def _leggi(chiave, nomi, files):
    if files is None:
        return _letti['esito'] if _letti['chiave'] == chiave else None
    fatture, scartati, letti = ([], [], [])
    for name, raw in zip(nomi, files):
        raw = bytes(raw)
        name = name or 'fattura.xml'
        if not _is_valid_name(name):
            motivo = 'estensione non valida'
        elif len(raw) > MAX_SIZE_BYTES:
            motivo = f'supera i {MAX_SIZE_MB} MB'
        elif len(raw) < 4:
            motivo = 'file vuoto'
        else:
            try:
                trovate = parse_fattura(raw, filename=name)
                fatture.extend(trovate)
                motivo = '' if trovate else 'nessun documento riconosciuto'
                if trovate:
                    letti.append((name, raw))
            except FormatoNonSupportato as e:
                motivo = str(e)
            except Exception:
                motivo = 'non leggibile'
        if motivo:
            scartati.append(f'{name} ({motivo})')
    _letti.update(chiave=chiave, esito=(fatture, scartati, letti))
    return (fatture, scartati, letti)

def _troppi(nomi, files):
    if files is not None and sum((len(f) for f in files)) > MAX_TOTAL_BYTES:
        raise Risposta(413, f'Troppi dati in un invio solo (massimo {MAX_TOTAL_MB} MB): dividi i file in più invii.')

def _leggi_ed_esporta(chiave, nomi, files):
    if len(nomi) > MAX_FILES:
        raise Risposta(400, f'Massimo {MAX_FILES} file per volta.')
    _troppi(nomi, files)
    letto = _leggi(chiave, nomi, files)
    if letto is None:
        return None
    fatture, scartati, letti = letto
    if not fatture:
        detail = 'Nessun dato estratto. I file inviati non contengono fatture riconoscibili.'
        if scartati:
            detail += f" File non elaborati: {', '.join(scartati)}."
        raise Risposta(422, detail)
    try:
        excel = export_excel_bytes(fatture)
    except Exception:
        raise Risposta(500, 'Errore nella generazione del file Excel.')
    return (fatture, scartati, letti, excel)

def _download(nomi, estensione, fatture, scartati):
    if len(nomi) > 1:
        fname = f'fatture_export.{estensione}'
    else:
        raw_name = (nomi[0] or 'fattura').rsplit('.', 1)[0].replace('.xml', '')
        safe_name = re.sub('[^a-zA-Z0-9._\\-]', '_', raw_name)[:50]
        fname = f'{safe_name}_fatture.{estensione}'
    return {'Content-Disposition': f"attachment; filename*=UTF-8''{urllib.parse.quote(fname, safe='._-')}", 'X-Fatture-Extracted': str(len(fatture)), 'X-Files-Skipped': urllib.parse.quote(json.dumps({'n': len(scartati), 'nomi': scartati[:20]}))}

def _unico(percorso, usati):
    radice, punto, est = percorso.rpartition('.')
    if not punto or '/' in est:
        radice, est = (percorso, '')
    candidato, n = (percorso, 2)
    while candidato.lower() in usati:
        candidato = f'{radice} ({n}).{est}' if est else f'{radice} ({n})'
        n += 1
    usati.add(candidato.lower())
    return candidato

def _risposta(fn):
    try:
        r = fn()
    except Risposta as e:
        return (json.dumps({'status': e.status, 'detail': e.detail}), b'')
    if r is None:
        return (json.dumps({'status': 0, 'rimanda': True}), b'')
    corpo, headers, tipo = r
    if isinstance(corpo, (bytes, bytearray)):
        return (json.dumps({'status': 200, 'headers': headers, 'tipo': tipo}), bytes(corpo))
    return (json.dumps({'status': 200, 'json': corpo}), b'')

def elabora(chiave, nomi, files):

    def fai():
        r = _leggi_ed_esporta(chiave, nomi, files)
        if r is None:
            return None
        fatture, scartati, _, excel = r
        return (excel, _download(nomi, 'xlsx', fatture, scartati), XLSX)
    return _risposta(fai)

def elabora_zip(chiave, nomi, files):

    def fai():
        r = _leggi_ed_esporta(chiave, nomi, files)
        if r is None:
            return None
        fatture, scartati, letti, excel = r
        buf = io.BytesIO()
        usati = {'fatture.xlsx'}
        with zipfile.ZipFile(buf, 'w', zipfile.ZIP_DEFLATED) as zf:
            zf.writestr('fatture.xlsx', excel)
            for nome, raw in letti:
                try:
                    contenuti = estrai_file(raw, nome)
                except Exception:
                    continue
                for percorso, dati in contenuti:
                    zf.writestr(_unico(percorso, usati), dati)
        return (buf.getvalue(), _download(nomi, 'zip', fatture, scartati), 'application/zip')
    return _risposta(fai)

def anteprima(chiave, nomi, files):

    def fai():
        if len(nomi) > MAX_FILES:
            raise Risposta(400, f'1-{MAX_FILES} file richiesti.')
        _troppi(nomi, files)
        letto = _leggi(chiave, nomi, files)
        if letto is None:
            return None
        fatture, scartati, _ = letto
        preview = [{'File': f['file'], 'Tipo': f['tipo'], 'Numero': f['numero'], 'Data': f['data'], 'Fornitore': f['fornitore']['nome'], 'Cliente': f['cliente']['nome'], 'Imponibile': round(f['tot_imponibile'], 2), 'Imposta': round(f['tot_imposta'], 2), 'Totale': f['totale_doc']} for f in fatture[:200]]
        righe = sum((len(f.get('righe') or []) for f in fatture))
        return ({'totale': len(fatture), 'anteprima': preview, 'troncato': len(fatture) > 200, 'scartati': scartati, 'righe': righe}, None, None)
    return _risposta(fai)
