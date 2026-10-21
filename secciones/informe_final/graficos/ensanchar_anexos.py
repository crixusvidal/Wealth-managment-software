"""Ensancha las tablas de los anexos de un Word del informe (10666 → 11506 twips, el ancho útil de la página) para que
no queden corridas a la izquierda con un hueco en el margen derecho.

Uso: python3 ensanchar_anexos.py entrada.docx salida.docx
"""
import re, shutil, sys, zipfile
from lxml import etree

W = "{http://schemas.openxmlformats.org/wordprocessingml/2006/main}"
ANCHO, VIEJO = 11506, 10666


def texto(e):
    return "".join(t.text or "" for t in e.iter(W + "t"))


def ensanchar(entrada, salida):
    z = zipfile.ZipFile(entrada)
    root = etree.fromstring(z.read("word/document.xml"))
    f, n, anexo = ANCHO / VIEJO, 0, False
    for ch in root.find(W + "body"):
        if ch.tag == W + "p" and re.match(r"(ANEXOS|APPENDICES): ", texto(ch).strip().upper()):
            anexo = True
        if not anexo or ch.tag != W + "tbl":
            continue
        if sum(int(g.get(W + "w")) for g in ch.find(W + "tblGrid")) != VIEJO:
            continue
        for g in ch.find(W + "tblGrid"):
            g.set(W + "w", str(round(int(g.get(W + "w")) * f)))
        for e in ch.iter(W + "tcW", W + "tblW"):
            if e.get(W + "type") == "dxa":
                e.set(W + "w", str(round(int(e.get(W + "w")) * f)))
        n += 1
    out = str(salida) + ".tmp"
    with zipfile.ZipFile(out, "w", zipfile.ZIP_DEFLATED) as zo:
        for it in z.infolist():
            zo.writestr(it, etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone=True)
                        if it.filename == "word/document.xml" else z.read(it.filename))
    shutil.move(out, salida)
    return n


if __name__ == "__main__":
    print(ensanchar(sys.argv[1], sys.argv[2]), "tablas ensanchadas")
