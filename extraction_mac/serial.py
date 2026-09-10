"""Extraction du numero de serie commun au Word et a l'Excel.

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
