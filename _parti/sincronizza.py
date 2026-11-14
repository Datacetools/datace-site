"""
Copia le parti comuni (_parti/) in tutte le pagine del sito.

Footer e intestazioni erano ricopiati a mano in ogni pagina: aggiungere o
togliere uno strumento voleva dire toccare 15-20 file, ed era facile
dimenticarne uno. Ora si modifica SOLO il file in _parti/ e si lancia, dalla
radice del sito:

    py _parti/sincronizza.py             # aggiorna le pagine
    py _parti/sincronizza.py --controlla # non scrive: dice quali pagine sono diverse

Le parti (brief 075 per le intestazioni):
- footer.html      fra <!-- DATACE FOOTER START --> e <!-- DATACE FOOTER END -->
- header.html      fra <!-- DATACE HEADER START --> e <!-- DATACE HEADER END -->
                   (sito: home, vetrine, guide, pagine di testo); la voce del
                   menu della sezione in cui si è (Guide, Chi sono) resta bianca
- header-sito-en.html, footer-en.html  fra <!-- DATACE HEADER EN START --> ... e <!-- DATACE FOOTER EN START --> ...:
                   solo la pagina inglese dei convertitori (en/excel-converters/), stessa struttura con le voci tradotte
- header-app.html  fra <!-- DATACE HEADER APP START --> e <!-- DATACE HEADER APP END -->
                   (F24, Fattura, Bonifici); nome e sottotitolo dello strumento
                   stanno in APP qui sotto. Fatture XML in bonifici ha la sua
                   intestazione fissa e piena (brief 061): non ha i marcatori.

Lo script sostituisce solo ciò che sta fra i marcatori: il resto della pagina
non lo tocca, e rispetta gli a-capo di ogni blocco. Le pagine senza marcatori
restano come sono. Le cartelle che iniziano con "_" non vengono pubblicate.
"""
import pathlib
import re
import sys

SITO = pathlib.Path(__file__).resolve().parent.parent

APP = {
    "f24-to-excel/app/index.html": ("F24 in Excel", "Estrai Debiti e Crediti dai tuoi F24"),
    "fattura-elettronica-excel/app/index.html": ("Fattura in Excel", "FPR12, FPA12, .p7m, notifiche SdI · Excel su 5 fogli"),
    "bonifici-massivi-xml/app/index.html": ("Bonifici massivi XML", "Da Excel a XML SEPA (CBI) · pronto per l'home banking"),
    "csv-in-excel/app/index.html": ("CSV in Excel", "Apri un CSV in Excel senza guai"),
    "json-in-excel/app/index.html": ("JSON in Excel", "Trasforma un JSON in una tabella Excel"),
}
# Voce del menu della sezione in cui si è: bianca e in grassetto, senza effetto al passaggio.
ATTIVA = {"guida": ("/guida/", "Guide"), "errori-fattura-elettronica": ("/guida/", "Guide"), "chi-siamo": ("/chi-siamo/", "Chi sono")}
VOCE = ('<a href="{href}" style="transition:color .2s" onmouseover="this.style.color=\'white\'" '
        'onmouseout="this.style.color=\'rgba(255,255,255,.62)\'">{testo}</a>')
VOCE_ATTIVA = '<a href="{href}" style="color:white;font-weight:600;">{testo}</a>'


def parte(nome):
    # utf-8-sig: se il file viene salvato con un BOM (lo fa PowerShell, alcuni
    # editor), il BOM NON deve finire nelle pagine: si accumulerebbe a ogni giro.
    return (SITO / "_parti" / nome).read_text(encoding="utf-8-sig").strip()


def header_sito(rel, testo):
    sezione = rel.split("/")[0]
    if sezione in ATTIVA:
        href, voce = ATTIVA[sezione]
        prima = VOCE.format(href=href, testo=voce)
        if prima not in testo:
            raise SystemExit(f"header.html: manca la voce «{voce}»")
        testo = testo.replace(prima, VOCE_ATTIVA.format(href=href, testo=voce))
    return testo


def header_app(rel, testo):
    if rel not in APP:
        raise SystemExit(f"{rel}: ha i marcatori dell'intestazione delle app ma non sta in APP")
    nome, sotto = APP[rel]
    return testo.replace("{{nome}}", nome).replace("{{sotto}}", sotto)


def footer_sito(rel, testo):
    # Solo la pagina italiana dei convertitori ha il link alla versione inglese, in fondo a «Risorse»
    # (brief 129); il footer comune a tutto il sito non cambia.
    if rel == "convertitori-in-excel/index.html":
        fine = '<a href="/#contatti" class="footer-link">Contatti</a>'
        if fine not in testo:
            raise SystemExit("footer.html: manca la voce «Contatti»")
        testo = testo.replace(fine, fine + '\n        <a href="/en/excel-converters/" class="footer-link" hreflang="en" lang="en">English</a>')
    return testo


PARTI = [
    ("FOOTER", parte("footer.html"), footer_sito),
    ("HEADER", parte("header.html"), header_sito),
    ("HEADER APP", parte("header-app.html"), header_app),
    # Pagina inglese dei convertitori: stessa intestazione e stesso piè, con le sole voci tradotte
    ("HEADER EN", parte("header-sito-en.html"), lambda rel, t: t),
    ("FOOTER EN", parte("footer-en.html"), lambda rel, t: t),
]


def pagine():
    for p in sorted(SITO.rglob("*.html")):
        rel = p.relative_to(SITO).parts
        if rel[0].startswith(("_", ".")) or rel[0] == "docs":
            continue
        yield p


def main(controlla: bool) -> int:
    diverse = []
    for p in pagine():
        rel = p.relative_to(SITO).as_posix()
        testo = p.read_bytes().decode("utf-8")
        nuovo = testo
        for nome, blocco, adatta in PARTI:
            cerca = re.compile(rf"<!-- DATACE {nome} START -->.*?<!-- DATACE {nome} END -->", re.S)
            trovato = cerca.search(nuovo)
            if not trovato:
                continue  # pagine senza questa parte (404, app con intestazione propria)
            # Gli a-capo del BLOCCO, non del file: alcune pagine li hanno misti, e
            # cambiarli farebbe risultare modificate righe identiche.
            nl = "\r\n" if "\r\n" in trovato.group(0) else "\n"
            sostituto = adatta(rel, blocco).replace("\n", nl)
            nuovo = cerca.sub(lambda _: sostituto, nuovo, count=1)
        if nuovo != testo:
            diverse.append(rel)
            if not controlla:
                p.write_bytes(nuovo.encode("utf-8"))
    azione = "da aggiornare" if controlla else "aggiornate"
    print(f"{len(diverse)} pagine {azione}", *diverse, sep="\n  ")
    return 1 if (controlla and diverse) else 0


if __name__ == "__main__":
    sys.exit(main("--controlla" in sys.argv))
