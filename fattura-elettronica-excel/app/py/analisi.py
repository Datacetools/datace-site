import re
from decimal import Decimal
from fattura_core import ESIGIBILITA, NATURA
SOMMATI = {'TD01', 'TD04', 'TD05', 'TD06', 'TD24', 'TD25'}
NON_DOCUMENTI = {'NOTIFICA', 'SDI-META'}
SPLIT = ESIGIBILITA['S']
QUOTA_TU = 0.6
PRIME = 10
CENT = Decimal('0.01')
_MESE = re.compile('^(\\d{4})-(\\d{2})')

def _d(x):
    return Decimal(str(x)) if x not in (None, '') else Decimal(0)

def _id(s):
    return s.get('piva') or s.get('cf') or ''

def _segno(f):
    if f.get('tipo_cod') == 'TD04' and _d(f.get('totale_doc')) > 0:
        return -1
    return 1

def _num(x):
    return float(x.quantize(CENT))

def _controparte(docs, tu, chi):
    mio = 'cliente' if chi == 'fornitore' else 'fornitore'
    somme, nomi, esclusi = ({}, {}, 0)
    for f, s, imp in docs:
        if tu and _id(f[mio]) != tu:
            esclusi += 1
            continue
        k = _id(f[chi]) or f[chi].get('nome') or ''
        somme[k] = somme.get(k, Decimal(0)) + s * imp
        nomi.setdefault(k, f[chi].get('nome') or '')
    ordine = sorted(somme, key=lambda k: (-abs(somme[k]), nomi[k]))
    altri = ordine[PRIME:]
    return {'voci': [[nomi[k], _num(somme[k])] for k in ordine[:PRIME]], 'altri': _num(sum((somme[k] for k in altri), Decimal(0))), 'n_altri': len(altri), 'esclusi': esclusi}

def calcola(fatture):
    letti = [f for f in fatture if f.get('tipo_cod') not in NON_DOCUMENTI]
    visti, unici, doppioni = (set(), [], 0)
    for f in letti:
        k = (_id(f['fornitore']) or f['fornitore'].get('nome') or '', (f.get('data') or '')[:4], f.get('numero') or '', f.get('tipo_cod') == 'TD04')
        if k in visti:
            doppioni += 1
            continue
        visti.add(k)
        unici.append(f)
    non_sommati = [f for f in unici if f.get('tipo_cod') not in SOMMATI]
    altra_valuta = [f for f in unici if f.get('tipo_cod') in SOMMATI and (f.get('divisa') or 'EUR') != 'EUR']
    sommati = [f for f in unici if f.get('tipo_cod') in SOMMATI and (f.get('divisa') or 'EUR') == 'EUR']
    imponibile = iva = totale = Decimal(0)
    mesi, aliquote, docs = ({}, {}, [])
    senza_data = note = split = reverse = 0
    for f in sommati:
        s = _segno(f)
        rie = f.get('riepilogo_iva') or []
        imp = sum((_d(r.get('imponibile')) for r in rie), Decimal(0)).quantize(CENT)
        tax = sum((_d(r.get('imposta')) for r in rie), Decimal(0)).quantize(CENT)
        imponibile += s * imp
        iva += s * tax
        totale += s * _d(f.get('totale_doc'))
        docs.append((f, s, imp))
        note += f.get('tipo_cod') == 'TD04'
        split += any((r.get('esigibilita') == SPLIT for r in rie))
        reverse += any(((r.get('natura_cod') or '').startswith('N6') for r in rie))
        m = _MESE.match(f.get('data') or '')
        if m and 1 <= int(m.group(2)) <= 12:
            mesi[m.group(0)] = mesi.get(m.group(0), Decimal(0)) + s * imp
        else:
            senza_data += 1
        for r in rie:
            a = _d(r.get('aliquota'))
            v = aliquote.setdefault((a, r.get('natura_cod') or ''), [Decimal(0), Decimal(0)])
            v[0] += s * _d(r.get('imponibile'))
            v[1] += s * _d(r.get('imposta'))
    conta, ruoli = ({}, {})
    for f in sommati:
        for ruolo in ('fornitore', 'cliente'):
            k = _id(f[ruolo])
            if k:
                ruoli.setdefault(k, {'fornitore': 0, 'cliente': 0})[ruolo] += 1
        for k in {_id(f['fornitore']), _id(f['cliente'])} - {''}:
            conta[k] = conta.get(k, 0) + 1
    tu, ruolo_tu, quota = ('', None, 0.0)
    if conta:
        tu = max(sorted(conta), key=lambda k: conta[k])
        quota = conta[tu] / len(sommati)
        ruolo_tu = 'cliente' if ruoli[tu]['cliente'] >= ruoli[tu]['fornitore'] else 'fornitore'
        if quota < QUOTA_TU:
            tu, ruolo_tu = ('', None)
    nome_tu = ''
    if tu:
        f = next((f for f in sommati if _id(f[ruolo_tu]) == tu))
        nome_tu = f[ruolo_tu].get('nome') or ''
    return {'letti': len(letti), 'documenti': len(sommati), 'imponibile': _num(imponibile), 'iva': _num(iva), 'totale': _num(totale), 'note_credito': note, 'split': split, 'reverse': reverse, 'doppioni': doppioni, 'altra_valuta': len(altra_valuta), 'non_sommati': len(non_sommati), 'senza_data': senza_data, 'tu': {'nome': nome_tu, 'quota': round(quota, 3), 'fatture': {'cliente': 'ricevute', 'fornitore': 'emesse'}.get(ruolo_tu, 'misti')}, 'ricevute': _controparte(docs, tu, 'fornitore'), 'emesse': _controparte(docs, tu, 'cliente'), 'mesi': [[k, _num(mesi[k])] for k in sorted(mesi)], 'aliquote': [[float(a), n, NATURA.get(n, n), _num(v[0]), _num(v[1])] for (a, n), v in sorted(aliquote.items())]}
