"""Lecture de l'adresse MAC dans un fichier Word.

Trois formats de fichier sont geres, sans dependance externe :

* .docx (et les .doc qui sont en realite des .docx renommes) : lecture directe
  du XML dans le zip, corps + en-tetes + pieds de page + zones de texte ;
* .doc binaire Word 97-2003 (conteneur OLE2) : balayage des octets en cp1252 et
  en UTF-16LE. Une adresse MAC est un motif tres caracteristique, ce balayage
  suffit sans embarquer un analyseur .doc complet ;
* .rtf renomme en .doc : decodage texte simple.

La recherche privilegie toujours une ligne portant le libelle « MAC ».
"""

from __future__ import annotations

import html
import re
import zipfile

# 00:30:D6:4C:6E:05 ou 00-30-D6-4C-6E-05 (meme separateur repete)
_MAC_SEPARE = re.compile(
    r"(?<![0-9A-Za-z])([0-9A-Fa-f]{2})([:\-])"
    r"(?:[0-9A-Fa-f]{2}\2){4}[0-9A-Fa-f]{2}(?![0-9A-Za-z])"
)
# 0030.D64C.6E05 (notation Cisco)
_MAC_POINTS = re.compile(
    r"(?<![0-9A-Za-z])[0-9A-Fa-f]{4}\.[0-9A-Fa-f]{4}\.[0-9A-Fa-f]{4}(?![0-9A-Za-z])"
)
# 0030D64C6E05 : accepte uniquement sur une ligne portant le libelle MAC,
# sinon le risque de confusion avec un numero de serie est trop grand.
_MAC_BRUT = re.compile(r"(?<![0-9A-Za-z])[0-9A-Fa-f]{12}(?![0-9A-Za-z])")

_LIBELLE_MAC = re.compile(r"\bMAC\b", re.IGNORECASE)

_MAC_INVALIDES = {"000000000000", "FFFFFFFFFFFF"}


def normaliser(brut, separateur=":", majuscules=True):
    """Transforme une MAC quelconque en ``xx:xx:xx:xx:xx:xx``."""
    hexa = re.sub(r"[^0-9A-Fa-f]", "", brut)
    hexa = hexa.upper() if majuscules else hexa.lower()
    return separateur.join(hexa[i:i + 2] for i in range(0, 12, 2))


def _candidats(fragment, autoriser_brut):
    """Retourne les MAC (12 caracteres hexa, majuscules) trouvees dans *fragment*."""
    trouves = []
    for regex in (_MAC_SEPARE, _MAC_POINTS):
        for match in regex.finditer(fragment):
            valeur = re.sub(r"[^0-9A-Fa-f]", "", match.group(0)).upper()
            if valeur not in trouves:
                trouves.append(valeur)
    if not trouves and autoriser_brut:
        for match in _MAC_BRUT.finditer(fragment):
            valeur = match.group(0).upper()
            if valeur not in trouves:
                trouves.append(valeur)
    return trouves


def chercher_mac(texte):
    """Retourne (mac_hexa, erreur) a partir du texte complet du document."""
    lignes = [ligne for ligne in texte.splitlines() if ligne.strip()]

    # 1) priorite absolue aux lignes portant le libelle « MAC »
    etiquetees = []
    for ligne in lignes:
        if _LIBELLE_MAC.search(ligne):
            for valeur in _candidats(ligne, autoriser_brut=True):
                if valeur not in etiquetees:
                    etiquetees.append(valeur)
    retenus = etiquetees

    # 2) a defaut, une MAC ecrite avec separateurs n'importe ou dans le document
    if not retenus:
        retenus = []
        for ligne in lignes:
            for valeur in _candidats(ligne, autoriser_brut=False):
                if valeur not in retenus:
                    retenus.append(valeur)

    if not retenus:
        return None, "aucune adresse MAC trouvée dans le document"
    if len(retenus) > 1:
        lisibles = ", ".join(normaliser(v) for v in retenus)
        return None, "plusieurs adresses MAC différentes trouvées : %s" % lisibles
    mac = retenus[0]
    if mac in _MAC_INVALIDES:
        return None, "adresse MAC invalide dans le document : %s" % normaliser(mac)
    return mac, None


# --------------------------------------------------------------------------
# lecture du texte selon le type reel du fichier
# --------------------------------------------------------------------------

def _xml_vers_texte(xml):
    xml = re.sub(r"<w:(?:tab|br)\b[^>]*/?>", "\n", xml)
    xml = re.sub(r"</w:p\s*>", "\n", xml)
    xml = re.sub(r"<[^>]+>", "", xml)
    return html.unescape(xml)


def _texte_docx(chemin):
    morceaux = []
    with zipfile.ZipFile(chemin) as zf:
        noms = [n for n in zf.namelist()
                if re.match(r"word/(document|header\d*|footer\d*|footnotes|endnotes)\.xml$", n)]
        # le corps d'abord : c'est la qu'on attend le libelle MAC
        noms.sort(key=lambda n: (n != "word/document.xml", n))
        for nom in noms:
            try:
                morceaux.append(_xml_vers_texte(zf.read(nom).decode("utf-8", "replace")))
            except KeyError:
                continue
    return "\n".join(morceaux)


def _nettoyer_controles(texte):
    return re.sub(r"[\x00-\x08\x0b-\x1f\x7f]+", "\n", texte).replace("\r", "\n")


def _texte_binaire(donnees):
    morceaux = [donnees.decode("cp1252", "replace")]
    for decalage in (0, 1):
        bloc = donnees[decalage:]
        bloc = bloc[: len(bloc) - (len(bloc) % 2)]
        morceaux.append(bloc.decode("utf-16-le", "replace"))
    return _nettoyer_controles("\n".join(morceaux))


def lire_texte(chemin):
    """Retourne le texte brut du document Word, quel que soit son format reel."""
    with open(chemin, "rb") as fichier:
        entete = fichier.read(8)
    if entete.startswith(b"PK\x03\x04"):
        return _texte_docx(chemin)
    with open(chemin, "rb") as fichier:
        donnees = fichier.read()
    if entete.startswith(b"\xd0\xcf\x11\xe0"):          # conteneur OLE2 (.doc)
        return _texte_binaire(donnees)
    return _nettoyer_controles(donnees.decode("cp1252", "replace"))  # rtf ou texte


def extraire_mac(chemin, separateur=":", majuscules=True):
    """Retourne (mac_formatee, erreur) pour un fichier Word."""
    try:
        texte = lire_texte(chemin)
    except zipfile.BadZipFile:
        return None, "fichier Word illisible (archive corrompue)"
    except OSError as erreur:
        return None, "fichier Word illisible : %s" % erreur
    mac, erreur = chercher_mac(texte)
    if erreur:
        return None, erreur
    return normaliser(mac, separateur, majuscules), None
