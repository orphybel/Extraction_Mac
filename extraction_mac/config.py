"""Profils de configuration, enregistres pour etre rappeles d'une fois sur l'autre.

Chaque profil retient le dossier, l'onglet et la cellule ou ecrire, ainsi que
les options d'ecriture.

Le fichier est ecrit **a cote du programme** : dans le dossier de
``ExtractionMAC.exe``, ou a la racine du projet quand on lance les sources.
L'outil reste ainsi portable — on copie le dossier, les profils suivent.

Si ce dossier n'est pas accessible en ecriture (executable pose dans
``C:\\Program Files``, sur un partage reseau en lecture seule, sur une cle
protegee...), la configuration bascule automatiquement vers
``%APPDATA%\\ExtractionMac``. Sans ce repli, l'enregistrement des profils
echouerait sans que l'utilisateur comprenne pourquoi.

A la lecture, le fichier place a cote du programme est prioritaire.
"""

from __future__ import annotations

import copy
import json
import os
import sys
import tempfile

NOM_FICHIER = "ExtractionMAC-config.json"

PROFIL_DEFAUT = {
    "dossier": "",
    "feuille": "Constit produit",
    "cellule": "F27",
    "sous_dossiers": False,
    "ecraser": False,
    "separateur_mac": ":",
    "mac_majuscules": True,
    "groupes_numero": [6, 7, 5],
}

CLES_ENTIERES = ("groupes_numero",)
CLES_BOOLEENNES = ("sous_dossiers", "ecraser", "mac_majuscules")


def dossier_programme():
    """Dossier de l'executable, ou racine du projet quand on lance les sources.

    Avec PyInstaller en un seul fichier, ``__file__`` pointe vers le dossier
    temporaire de decompression : c'est ``sys.executable`` qui donne l'endroit
    ou l'utilisateur a reellement pose le programme.
    """
    if getattr(sys, "frozen", False):
        return os.path.dirname(os.path.abspath(sys.executable))
    return os.path.dirname(os.path.dirname(os.path.abspath(__file__)))


def dossier_repli():
    """Emplacement utilise quand on ne peut pas ecrire a cote du programme."""
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, "ExtractionMac")
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "extraction_mac")


def _accessible_en_ecriture(dossier):
    """Teste par une ecriture reelle : sous Windows, os.access est peu fiable."""
    try:
        os.makedirs(dossier, exist_ok=True)
        descripteur, temoin = tempfile.mkstemp(prefix=".ecriture-", dir=dossier)
        os.close(descripteur)
        os.remove(temoin)
        return True
    except OSError:
        return False


def dossier_config():
    """Dossier ou la configuration sera ecrite."""
    programme = dossier_programme()
    if _accessible_en_ecriture(programme):
        return programme
    return dossier_repli()


def chemin_config():
    return os.path.join(dossier_config(), NOM_FICHIER)


def chemins_lecture():
    """Emplacements consultes a la lecture, du plus prioritaire au moins prioritaire."""
    chemins = []
    for dossier in (dossier_programme(), dossier_repli()):
        chemin = os.path.join(dossier, NOM_FICHIER)
        if chemin not in chemins:
            chemins.append(chemin)
    return chemins


def profil_vide():
    return copy.deepcopy(PROFIL_DEFAUT)


def normaliser_profil(brut):
    """Complete un profil lu sur disque avec les valeurs par defaut manquantes."""
    profil = profil_vide()
    if isinstance(brut, dict):
        for cle, valeur in brut.items():
            if cle not in profil:
                continue
            if cle in CLES_BOOLEENNES:
                profil[cle] = bool(valeur)
            elif cle in CLES_ENTIERES:
                try:
                    groupes = [int(n) for n in valeur]
                except (TypeError, ValueError):
                    continue
                if groupes and all(n > 0 for n in groupes):
                    profil[cle] = groupes
            else:
                profil[cle] = "" if valeur is None else str(valeur)
    return profil


def charger():
    """Retourne (profils, dernier_profil).

    Le premier emplacement lisible gagne, celui d'a cote du programme d'abord.
    Ne leve jamais : un fichier illisible est traite comme une configuration vide.
    """
    for chemin in chemins_lecture():
        try:
            with open(chemin, encoding="utf-8") as fichier:
                donnees = json.load(fichier)
        except (OSError, ValueError):
            continue
        if not isinstance(donnees, dict):
            continue
        bruts = donnees.get("profils") or donnees.get("profiles") or {}
        profils = {str(nom): normaliser_profil(valeur)
                   for nom, valeur in bruts.items() if isinstance(nom, str)}
        dernier = donnees.get("dernier_profil") or ""
        return profils, (dernier if dernier in profils else "")
    return {}, ""


def enregistrer(profils, dernier_profil=""):
    """Ecrit la configuration de facon atomique et retourne le chemin utilise.

    Leve OSError si aucun des deux emplacements n'est accessible.
    """
    dossier = dossier_config()
    os.makedirs(dossier, exist_ok=True)
    destination = os.path.join(dossier, NOM_FICHIER)
    donnees = {
        "version": 1,
        "dernier_profil": dernier_profil,
        "profils": {nom: normaliser_profil(profil) for nom, profil in profils.items()},
    }
    descripteur, temporaire = tempfile.mkstemp(suffix=".tmp", dir=dossier)
    try:
        with os.fdopen(descripteur, "w", encoding="utf-8") as fichier:
            json.dump(donnees, fichier, indent=2, ensure_ascii=False)
        os.replace(temporaire, destination)
    except BaseException:
        if os.path.exists(temporaire):
            os.remove(temporaire)
        raise
    return destination
