import io
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser
from pdfminer.pdftypes import resolve1

def conta(files):
    n = q = 0
    for raw in files:
        try:
            doc = PDFDocument(PDFParser(io.BytesIO(bytes(raw))))
            p = sum((1 for _ in PDFPage.create_pages(doc)))
            n += p
            info = resolve1(doc.info[0]) if doc.info else {}
            if resolve1(info.get('Author')) == b'Agenzia Entrate':
                q += p
        except Exception:
            pass
    return (n, q)
