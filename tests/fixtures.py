"""Fabrique des .docx et .xlsx minimaux, pour tester sans donnee client."""

from __future__ import annotations

import zipfile

_CT_DOCX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

_RELS_DOCX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""


def ecrire_docx(chemin, paragraphes):
    """Cree un .docx dont chaque paragraphe est coupe en deux « runs »,
    comme le fait Word : cela verifie que le libelle et la valeur sont
    bien recolles avant la recherche."""
    corps = []
    for texte in paragraphes:
        milieu = len(texte) // 2
        corps.append(
            "<w:p><w:r><w:t xml:space=\"preserve\">%s</w:t></w:r>"
            "<w:r><w:t xml:space=\"preserve\">%s</w:t></w:r></w:p>"
            % (texte[:milieu].replace("&", "&amp;"), texte[milieu:].replace("&", "&amp;")))
    document = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">'
        "<w:body>%s</w:body></w:document>" % "".join(corps))
    with zipfile.ZipFile(chemin, "w") as zf:
        zf.writestr("[Content_Types].xml", _CT_DOCX)
        zf.writestr("_rels/.rels", _RELS_DOCX)
        zf.writestr("word/document.xml", document)
    return chemin


_CT_XLSX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
<Default Extension="xml" ContentType="application/xml"/>
<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>
<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>
</Types>"""

_RELS_XLSX = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
</Relationships>"""

_RELS_WB = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>
<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>
</Relationships>"""


def ecrire_xlsx(chemin, nom_feuille="Constit produit", lignes=None, fusions=(), piece_annexe=True):
    """*lignes* : {numero_ligne: [(ref, index_chaine_partagee | None, style), ...]}."""
    lignes = lignes if lignes is not None else {27: [("B27", 0, "172"), ("F27", None, "98")]}
    chaines = ["@MAC", "N° Série"]
    corps = []
    for numero in sorted(lignes):
        cellules = []
        for ref, index, style in lignes[numero]:
            attribut = ' s="%s"' % style if style else ""
            if index is None:
                cellules.append('<c r="%s"%s/>' % (ref, attribut))
            else:
                cellules.append('<c r="%s"%s t="s"><v>%d</v></c>' % (ref, attribut, index))
        corps.append('<row r="%d" spans="1:7">%s</row>' % (numero, "".join(cellules)))
    bloc_fusions = ""
    if fusions:
        bloc_fusions = '<mergeCells count="%d">%s</mergeCells>' % (
            len(fusions), "".join('<mergeCell ref="%s"/>' % f for f in fusions))
    feuille = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"'
        ' mc:Ignorable="x14ac"'
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">'
        '<dimension ref="A1:G40"/><sheetData>%s</sheetData>%s'
        '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        "</worksheet>" % ("".join(corps), bloc_fusions))
    workbook = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"'
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
        '<sheets><sheet name="%s" sheetId="1" r:id="rId1"/></sheets></workbook>'
        % nom_feuille.replace("&", "&amp;"))
    shared = (
        '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
        '<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="%d" uniqueCount="%d">%s</sst>'
        % (len(chaines), len(chaines),
           "".join("<si><t>%s</t></si>" % c for c in chaines)))
    with zipfile.ZipFile(chemin, "w") as zf:
        zf.writestr("[Content_Types].xml", _CT_XLSX)
        zf.writestr("_rels/.rels", _RELS_XLSX)
        zf.writestr("xl/workbook.xml", workbook)
        zf.writestr("xl/_rels/workbook.xml.rels", _RELS_WB)
        zf.writestr("xl/worksheets/sheet1.xml", feuille)
        zf.writestr("xl/sharedStrings.xml", shared)
        if piece_annexe:                      # imite un dessin / controle a preserver
            zf.writestr(
                "xl/drawings/drawing1.xml",
                '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
                '<xdr:wsDr xmlns:xdr="http://schemas.openxmlformats.org/'
                'drawingml/2006/spreadsheetDrawing"/>')
    return chemin
