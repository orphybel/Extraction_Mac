"""Le numero de serie ``xxxxxx-yyyyyyy-zzzzz``, lu et compose.

Le numero a toujours la forme ``xxxxxx-yyyyyyy-zzzzz`` (6 / 7 / 5 chiffres par
defaut). Il sert de cle d'appairage entre un .doc et son .xlsx.

Les separateurs sont tolerants : ``260918-0144215-00092``,
``260918_0144215_00092`` ou meme ``260918014421500092`` donnent la meme cle
canonique ``260918-0144215-00092``.
"""

from __future__ import annotations

import re

DEFAULT_GROUPS = (6, 7, 5)


def _pattern(groups):
    """Construit le motif : groupes de chiffres separes par au plus un caractere non alphanumerique.

    Les garde-fous ``(?<![0-9])`` / ``(?![0-9])`` empechent de tomber au milieu
    d'une suite de chiffres plus longue et d'en extraire une cle fantaisiste.
    """
    body = "[^0-9A-Za-z]?".join("([0-9]{%d})" % n for n in groups)
    return re.compile("(?<![0-9])" + body + "(?![0-9])")


def canonical(groups_values):
    """Reassemble les groupes sous la forme canonique avec des tirets."""
    return "-".join(groups_values)


def find_keys(text, groups=DEFAULT_GROUPS):
    """Retourne la liste des cles distinctes trouvees dans *text*, dans l'ordre."""
    found = []
    for match in _pattern(tuple(groups)).finditer(text or ""):
        key = canonical(match.groups())
        if key not in found:
            found.append(key)
    return found


def key_from_filename(name, groups=DEFAULT_GROUPS):
    """Retourne (cle, erreur). La cle vaut None si absente ou ambigue."""
    stem = re.sub(r"\.[A-Za-z0-9]{1,5}$", "", name)
    keys = find_keys(stem, groups)
    if not keys:
        attendu = "-".join("x" * n for n in groups)
        return None, "aucun numéro de série (%s) dans le nom du fichier" % attendu
    if len(keys) > 1:
        return None, "plusieurs numéros de série dans le nom du fichier : %s" % ", ".join(keys)
    return keys[0], None


# --------------------------------------------------------------------------
# composition d'un numero a partir d'une saisie partielle (releve au banc)
# --------------------------------------------------------------------------

def forme_attendue(groups=DEFAULT_GROUPS):
    return "-".join("x" * n for n in groups)


def normaliser_prefixe(prefixe, groups=DEFAULT_GROUPS):
    """Retourne (prefixe_canonique, erreur).

    Accepte ``260918-0144215``, ``260918 0144215``, ``2609180144215`` ou meme un
    numero complet colle depuis un nom de fichier, dont seuls les premiers
    groupes sont retenus.
    """
    groups = tuple(groups)
    if len(groups) < 2:
        return None, "le numéro doit comporter au moins deux groupes"
    tete = groups[:-1]
    chiffres = re.sub(r"[^0-9]", "", prefixe or "")
    if not chiffres:
        return None, "préfixe vide ; attendu %s" % forme_attendue(tete)
    attendu = sum(tete)
    if len(chiffres) < attendu:
        return None, ("préfixe trop court : %d chiffres au lieu de %d (attendu %s)"
                      % (len(chiffres), attendu, forme_attendue(tete)))
    chiffres = chiffres[:attendu]          # un numero complet colle par megarde
    morceaux, position = [], 0
    for longueur in tete:
        morceaux.append(chiffres[position:position + longueur])
        position += longueur
    return "-".join(morceaux), None


def composer(prefixe, saisie, groups=DEFAULT_GROUPS):
    """Retourne (numero_complet, erreur) a partir du prefixe et du dernier groupe.

    La saisie est completee a gauche par des zeros : ``92`` devient ``00092``.
    C'est tout l'interet au banc — trois caracteres tapes au lieu de vingt.
    Un numero complet colle dans le champ est accepte tel quel, pour ne pas
    bloquer l'operateur qui recopie depuis un nom de fichier.
    """
    groups = tuple(groups)
    brut = (saisie or "").strip()
    if not brut:
        return None, "numéro d'appareil vide"

    complets = find_keys(brut, groups)
    if len(complets) == 1 and len(re.sub(r"[^0-9]", "", brut)) == sum(groups):
        return complets[0], None

    if re.search(r"[^0-9\s]", brut):
        return None, "le numéro d'appareil ne doit contenir que des chiffres"
    chiffres = re.sub(r"\s", "", brut)
    dernier = groups[-1]
    if len(chiffres) > dernier:
        return None, ("numéro trop long : %d chiffres pour un groupe final de %d"
                      % (len(chiffres), dernier))

    tete, erreur = normaliser_prefixe(prefixe, groups)
    if erreur:
        return None, erreur
    return "%s-%s" % (tete, chiffres.zfill(dernier)), None


def partie_finale(numero):
    """Dernier groupe d'un numero complet, pour l'affichage."""
    return (numero or "").rsplit("-", 1)[-1]


def deduire_prefixe(noms, groups=DEFAULT_GROUPS):
    """Retourne (prefixe, message) deduit d'une liste de noms de fichiers.

    Sert a pre-remplir le champ « Prefixe » a partir des fiches Excel presentes
    dans le dossier : l'operateur n'a alors plus rien a regler. La deduction
    n'aboutit que si **tous** les numeros trouves partagent la meme tete ; sinon
    on prefere ne rien proposer plutot qu'une valeur fausse.
    """
    groups = tuple(groups)
    tetes, trouves = [], 0
    for nom in noms:
        numero, _ = key_from_filename(nom, groups)
        if not numero:
            continue
        tete = numero.rsplit("-", 1)[0]
        if tete not in tetes:
            tetes.append(tete)
        trouves += 1
    if not tetes:
        return None, "aucun numéro de série reconnu dans les noms de fichiers"
    if len(tetes) > 1:
        return None, ("plusieurs préfixes différents dans le dossier : %s"
                      % ", ".join(tetes[:4]))
    return tetes[0], "préfixe déduit de %d fichier(s) : %s" % (trouves, tetes[0])
