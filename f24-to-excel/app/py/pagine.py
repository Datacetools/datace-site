import io
from pdfminer.pdfdocument import PDFDocument
from pdfminer.pdfpage import PDFPage
from pdfminer.pdfparser import PDFParser

def conta(files):
    n = 0
    for raw in files:
        try:
            n += sum((1 for _ in PDFPage.create_pages(PDFDocument(PDFParser(io.BytesIO(bytes(raw)))))))
        except Exception:
            pass
    return n
