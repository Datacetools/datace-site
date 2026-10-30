"""
Rigenera images/og-datace.png (1200×630, l'anteprima di tutte le pagine) da
_og/og-datace.html, fotografata con Edge senza finestra. Brief 049.

Quando cambia uno strumento: si cambia la lista in _og/og-datace.html e, dalla radice
del sito, si lancia

    py _og/genera.py

Poi si guarda l'immagine (anche a 300 px di larghezza, come in un feed da telefono).
Nessun servizio esterno: Edge è quello installato su Windows. Le cartelle che iniziano
con "_" non vengono pubblicate da GitHub Pages.
"""
import pathlib
import struct
import subprocess
import sys
import tempfile

SITO = pathlib.Path(__file__).resolve().parent.parent
PAGINA = SITO / "_og" / "og-datace.html"
PNG = SITO / "images" / "og-datace.png"
EDGE = pathlib.Path(r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")

with tempfile.TemporaryDirectory() as profilo:   # profilo nuovo: niente estensioni né cache
    subprocess.run([str(EDGE), "--headless", "--disable-gpu", "--hide-scrollbars", "--force-device-scale-factor=1",
                    f"--user-data-dir={profilo}", "--window-size=1200,630", "--virtual-time-budget=3000",
                    f"--screenshot={PNG}", PAGINA.as_uri()], check=True, capture_output=True, timeout=120)

larghezza, altezza = struct.unpack(">II", PNG.read_bytes()[16:24])   # intestazione PNG
if (larghezza, altezza) != (1200, 630):
    sys.exit(f"misura sbagliata: {larghezza}x{altezza} invece di 1200x630")
print(f"scritta {PNG.relative_to(SITO)}: {larghezza}x{altezza}, {PNG.stat().st_size // 1024} KB")
