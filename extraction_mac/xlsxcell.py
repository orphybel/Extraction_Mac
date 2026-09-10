"""Lecture et ecriture d'une cellule dans un .xlsx / .xlsm, sans dependance.

L'ecriture est chirurgicale : seule la feuille concernee est modifiee, toutes
les autres pieces du zip sont recopiees octet pour octet. C'est volontaire.
Une reecriture complete par une bibliotheque tierce ferait perdre les dessins,
les controles de formulaire, les reglages d'impression et les mises en forme
conditionnelles etendues que contiennent ces fiches de test.
"""

from __future__ import annotations

import html
import os
import re
import shutil
import tempfile
import zipfile

NS_REL = "http://schemas.openxmlformats.org/package/2006/relationships"

_REF = re.compile(r"^([A-Za-z]{1,3})([1-9][0-9]*)$")


class ErreurExcel(Exception):
    """Probleme rencontre sur un classeur (feuille absente, format non gere...)."""


# --------------------------------------------------------------------------
# references de cellules
# --------------------------------------------------------------------------

def decouper_ref(ref):
    """'F27' -> ('F', 27). Leve ErreurExcel si la reference est invalide."""
    match = _REF.match((ref or "").strip().replace("$", ""))
    if not match:
        raise ErreurExcel("référence de cellule invalide : %r (attendu par exemple F27)" % ref)
    return match.group(1).upper(), int(match.group(2))


def col_vers_index(lettres):
    index = 0
    for caractere in lettres.upper():
        index = index * 26 + (ord(caractere) - 64)
    return index


def normaliser_ref(ref):
    lettres, ligne = decouper_ref(ref)
    return "%s%d" % (lettres, ligne)


# --------------------------------------------------------------------------
# lecture de la structure du classeur
# --------------------------------------------------------------------------

def _lire(zf, nom):
    return zf.read(nom).decode("utf-8", "replace")


def _feuilles(zf):
    """Retourne [(nom_feuille, piece_xml), ...] dans l'ordre des onglets."""
    workbook = _lire(zf, "xl/workbook.xml")
    liens = {}
    for match in re.finditer(r"<Relationship\b[^>]*>", _lire(zf, "xl/_rels/workbook.xml.rels")):
        balise = match.group(0)
        ident = re.search(r'Id="([^"]+)"', balise)
        cible = re.search(r'Target="([^"]+)"', balise)
        if ident and cible:
            chemin = cible.group(1)
            if chemin.startswith("/"):
                chemin = chemin.lstrip("/")
            elif not chemin.startswith("xl/"):
                chemin = "xl/" + chemin.lstrip("./")
            liens[ident.group(1)] = chemin
    feuilles = []
    for match in re.finditer(r"<sheet\b[^>]*/>", workbook):
        balise = match.group(0)
        nom = re.search(r'\sname="([^"]*)"', balise)
        ident = re.search(r'r:id="([^"]+)"', balise)
        if nom and ident and ident.group(1) in liens:
            feuilles.append((html.unescape(nom.group(1)), liens[ident.group(1)]))
    return feuilles


def lister_feuilles(chemin):
    """Retourne les noms d'onglets du classeur, dans l'ordre."""
    _verifier_format(chemin)
    with zipfile.ZipFile(chemin) as zf:
        return [nom for nom, _ in _feuilles(zf)]


def _verifier_format(chemin):
    extension = os.path.splitext(chemin)[1].lower()
    if extension in (".xls", ".xlt"):
        raise ErreurExcel("format Excel 97-2003 (.xls) non géré ; enregistrer le fichier en .xlsx")
    if not zipfile.is_zipfile(chemin):
        raise ErreurExcel("fichier Excel illisible ou corrompu")


def _piece_feuille(zf, nom_feuille):
    feuilles = _feuilles(zf)
    for nom, piece in feuilles:
        if nom == nom_feuille:
            return piece
    for nom, piece in feuilles:                      # tolerance casse / espaces
        if nom.strip().casefold() == (nom_feuille or "").strip().casefold():
            return piece
    raise ErreurExcel("feuille %r absente du classeur (onglets : %s)"
                      % (nom_feuille, ", ".join(n for n, _ in feuilles)))


def _chaines_partagees(zf):
    try:
        xml = _lire(zf, "xl/sharedStrings.xml")
    except KeyError:
        return []
    chaines = []
    for match in re.finditer(r"<si\b[^>]*>(.*?)</si>|<si\b[^>]*/>", xml, re.S):
        corps = match.group(1) or ""
        textes = re.findall(r"<t\b[^>]*>(.*?)</t>", corps, re.S)
        chaines.append(html.unescape("".join(textes)))
    return chaines


# --------------------------------------------------------------------------
# balayage du XML de feuille (les litteraux </c>, </row>, </sheetData> ne
# peuvent pas apparaitre dans du texte XML : '<' y est toujours echappe)
# --------------------------------------------------------------------------

def _elements(xml, balise, debut=0, fin=None):
    """Retourne [(debut, fin, balise_ouvrante), ...] pour chaque element *balise*."""
    fin = len(xml) if fin is None else fin
    ouverture = re.compile(r"<%s\b[^>]*>" % balise)
    resultat = []
    position = debut
    while True:
        match = ouverture.search(xml, position, fin)
        if not match:
            return resultat
        if match.group(0).endswith("/>"):
            resultat.append((match.start(), match.end(), match.group(0)))
            position = match.end()
            continue
        cloture = xml.find("</%s>" % balise, match.end(), fin)
        if cloture == -1:
            return resultat
        cloture += len(balise) + 3
        resultat.append((match.start(), cloture, match.group(0)))
        position = cloture


def _attribut(balise, nom):
    match = re.search(r'\s%s="([^"]*)"' % nom, balise)
    return match.group(1) if match else None


def _zone_sheet_data(xml):
    """Retourne (xml, debut_contenu, fin_contenu) en depliant un <sheetData/> vide."""
    match = re.search(r"<sheetData\b[^>]*?/>", xml)
    if match:
        xml = xml[:match.start()] + "<sheetData></sheetData>" + xml[match.end():]
    ouverture = re.search(r"<sheetData\b[^>]*>", xml)
    if not ouverture:
        raise ErreurExcel("feuille Excel illisible (<sheetData> introuvable)")
    fermeture = xml.find("</sheetData>", ouverture.end())
    if fermeture == -1:
        raise ErreurExcel("feuille Excel illisible (</sheetData> introuvable)")
    return xml, ouverture.end(), fermeture


def _fusions(xml):
    """Retourne [((col1, lig1), (col2, lig2), 'ancre'), ...] des cellules fusionnees."""
    plages = []
    for match in re.finditer(r'<mergeCell\b[^>]*ref="([^"]+)"', xml):
        ref = match.group(1)
        if ":" not in ref:
            continue
        gauche, droite = ref.split(":", 1)
        try:
            l1, r1 = decouper_ref(gauche)
            l2, r2 = decouper_ref(droite)
        except ErreurExcel:
            continue
        plages.append(((col_vers_index(l1), r1), (col_vers_index(l2), r2), "%s%d" % (l1, r1)))
    return plages


def resoudre_fusion(xml, ref):
    """Si *ref* est dans une plage fusionnee, retourne la cellule d'ancrage."""
    lettres, ligne = decouper_ref(ref)
    colonne = col_vers_index(lettres)
    for (c1, l1), (c2, l2), ancre in _fusions(xml):
        if c1 <= colonne <= c2 and l1 <= ligne <= l2:
            return ancre
    return "%s%d" % (lettres, ligne)


def _texte_cellule(balise_ouvrante, contenu, chaines):
    type_cellule = _attribut(balise_ouvrante, "t") or "n"
    if type_cellule == "inlineStr":
        return html.unescape("".join(re.findall(r"<t\b[^>]*>(.*?)</t>", contenu, re.S))) or None
    valeur = re.search(r"<v\b[^>]*>(.*?)</v>", contenu, re.S)
    if not valeur:
        return None
    brut = html.unescape(valeur.group(1))
    if type_cellule == "s":
        try:
            return chaines[int(brut)]
        except (ValueError, IndexError):
            return None
    return brut


def lire_cellule(chemin, nom_feuille, ref):
    """Retourne le contenu texte de la cellule, ou None si elle est vide."""
    _verifier_format(chemin)
    with zipfile.ZipFile(chemin) as zf:
        xml = _lire(zf, _piece_feuille(zf, nom_feuille))
        chaines = _chaines_partagees(zf)
    cible = resoudre_fusion(xml, ref)
    xml, debut, fin = _zone_sheet_data(xml)
    _, ligne_cible = decouper_ref(cible)
    for d_ligne, f_ligne, balise in _elements(xml, "row", debut, fin):
        if _attribut(balise, "r") != str(ligne_cible):
            continue
        for d_cell, f_cell, balise_cell in _elements(xml, "c", d_ligne, f_ligne):
            if _attribut(balise_cell, "r") == cible:
                if balise_cell.endswith("/>"):
                    return None
                contenu = xml[d_cell + len(balise_cell):f_cell - 4]
                valeur = _texte_cellule(balise_cell, contenu, chaines)
                return valeur.strip() if isinstance(valeur, str) and valeur.strip() else valeur
    return None


def trouver_libelles(chemin, motif="MAC"):
    """Retourne [(feuille, cellule, texte), ...] des cellules contenant *motif*.

    Sert a retrouver ou se trouve le libelle @MAC quand la feuille et la
    cellule changent d'un modele de fiche a l'autre.
    """
    _verifier_format(chemin)
    recherche = (motif or "").casefold()
    resultats = []
    with zipfile.ZipFile(chemin) as zf:
        chaines = _chaines_partagees(zf)
        for nom_feuille, piece in _feuilles(zf):
            try:
                xml = _lire(zf, piece)
            except KeyError:
                continue
            _, debut, fin = _zone_sheet_data(xml)
            for d_ligne, f_ligne, _ in _elements(xml, "row", debut, fin):
                for d_cell, f_cell, balise in _elements(xml, "c", d_ligne, f_ligne):
                    if balise.endswith("/>"):
                        continue
                    texte = _texte_cellule(balise, xml[d_cell + len(balise):f_cell - 4], chaines)
                    if texte and recherche in texte.casefold():
                        resultats.append((nom_feuille, _attribut(balise, "r") or "?",
                                          " ".join(texte.split())))
    return resultats


# --------------------------------------------------------------------------
# ecriture
# --------------------------------------------------------------------------

def _sans_calcchain(contenu):
    """Retire la declaration de calcChain.xml quand la piece n'est plus recopiee."""
    texte = contenu.decode("utf-8", "replace")
    texte = re.sub(r'<Override\b[^>]*PartName="/xl/calcChain\.xml"[^>]*/>', "", texte)
    return texte.encode("utf-8")


def _echapper(valeur):
    return valeur.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _nouvelle_cellule(ref, valeur, style):
    attribut_style = ' s="%s"' % style if style else ""
    return ('<c r="%s"%s t="inlineStr"><is><t xml:space="preserve">%s</t></is></c>'
            % (ref, attribut_style, _echapper(valeur)))


def poser_valeur(xml, ref, valeur):
    """Retourne (nouveau_xml, formule_supprimee) apres ecriture de *valeur* en *ref*."""
    ref = resoudre_fusion(xml, ref)
    lettres, ligne_cible = decouper_ref(ref)
    colonne_cible = col_vers_index(lettres)
    xml, debut, fin = _zone_sheet_data(xml)

    lignes = _elements(xml, "row", debut, fin)
    existante = None
    insertion_ligne = fin
    for d_ligne, f_ligne, balise in lignes:
        try:
            numero = int(_attribut(balise, "r") or 0)
        except ValueError:
            continue
        if numero == ligne_cible:
            existante = (d_ligne, f_ligne, balise)
            break
        if numero > ligne_cible:
            insertion_ligne = d_ligne
            break

    if existante is None:
        bloc = '<row r="%d">%s</row>' % (ligne_cible, _nouvelle_cellule(ref, valeur, None))
        return xml[:insertion_ligne] + bloc + xml[insertion_ligne:], False

    d_ligne, f_ligne, balise_ligne = existante
    if balise_ligne.endswith("/>"):                  # <row r="27"/> : ligne vide
        bloc = (balise_ligne[:-2].rstrip() + ">" + _nouvelle_cellule(ref, valeur, None) + "</row>")
        return xml[:d_ligne] + bloc + xml[f_ligne:], False

    cellules = _elements(xml, "c", d_ligne, f_ligne)
    cible = None
    insertion_cellule = f_ligne - len("</row>")
    for d_cell, f_cell, balise_cell in cellules:
        reference = _attribut(balise_cell, "r") or ""
        match = _REF.match(reference)
        if not match:
            continue
        colonne = col_vers_index(match.group(1))
        if reference == ref:
            cible = (d_cell, f_cell, balise_cell)
            break
        if colonne > colonne_cible:
            insertion_cellule = d_cell
            break

    if cible is None:
        bloc = _nouvelle_cellule(ref, valeur, None)
        nouveau = xml[:insertion_cellule] + bloc + xml[insertion_cellule:]
        return _elargir_spans(nouveau, d_ligne, colonne_cible), False

    d_cell, f_cell, balise_cell = cible
    style = _attribut(balise_cell, "s")
    ancien = "" if balise_cell.endswith("/>") else xml[d_cell + len(balise_cell):f_cell - 4]
    formule_supprimee = "<f" in ancien
    bloc = _nouvelle_cellule(ref, valeur, style)
    return xml[:d_cell] + bloc + xml[f_cell:], formule_supprimee


def _elargir_spans(xml, debut_ligne, colonne):
    """Retire l'attribut spans devenu faux (il est facultatif, Excel le recalcule)."""
    fin_balise = xml.find(">", debut_ligne)
    if fin_balise == -1:
        return xml
    balise = xml[debut_ligne:fin_balise + 1]
    spans = _attribut(balise, "spans")
    if not spans or ":" not in spans:
        return xml
    try:
        bas, haut = (int(part) for part in spans.split(":", 1))
    except ValueError:
        return xml
    if bas <= colonne <= haut:
        return xml
    remplacee = re.sub(r'\sspans="[^"]*"', "", balise, count=1)
    return xml[:debut_ligne] + remplacee + xml[fin_balise + 1:]


def ecrire_cellule(chemin, nom_feuille, ref, valeur):
    """Ecrit *valeur* (texte) dans la cellule, en place, sans toucher au reste du classeur."""
    _verifier_format(chemin)
    with zipfile.ZipFile(chemin) as source:
        piece = _piece_feuille(source, nom_feuille)
        xml = _lire(source, piece)
        nouveau_xml, formule_supprimee = poser_valeur(xml, ref, valeur)
        if nouveau_xml == xml:
            return False
        donnees = nouveau_xml.encode("utf-8")
        dossier = os.path.dirname(os.path.abspath(chemin)) or "."
        descripteur, temporaire = tempfile.mkstemp(suffix=".xlsx.tmp", dir=dossier)
        os.close(descripteur)
        try:
            with zipfile.ZipFile(temporaire, "w") as cible:
                cible.comment = source.comment
                for info in source.infolist():
                    if formule_supprimee and info.filename == "xl/calcChain.xml":
                        continue          # Excel le reconstruit au prochain calcul
                    if info.filename == piece:
                        contenu = donnees
                    elif formule_supprimee and info.filename == "[Content_Types].xml":
                        contenu = _sans_calcchain(source.read(info.filename))
                    else:
                        contenu = source.read(info.filename)
                    copie = zipfile.ZipInfo(info.filename, date_time=info.date_time)
                    copie.compress_type = info.compress_type
                    copie.external_attr = info.external_attr
                    copie.internal_attr = info.internal_attr
                    copie.create_system = info.create_system
                    cible.writestr(copie, contenu)
        except BaseException:
            if os.path.exists(temporaire):
                os.remove(temporaire)
            raise
    shutil.copymode(chemin, temporaire)
    os.replace(temporaire, chemin)
    return True
