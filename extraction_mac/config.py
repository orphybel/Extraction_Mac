"""Profils de configuration, enregistres pour etre rappeles d'une fois sur l'autre.

Chaque profil retient le dossier, l'onglet et la cellule ou ecrire, ainsi que
les options d'ecriture. Sous Windows le fichier est place dans
``%APPDATA%\\ExtractionMac\\config.json``.
"""

from __future__ import annotations

import copy
import json
import os
import tempfile

NOM_FICHIER = "config.json"

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


def dossier_config():
    if os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
        return os.path.join(base, "ExtractionMac")
    base = os.environ.get("XDG_CONFIG_HOME") or os.path.join(os.path.expanduser("~"), ".config")
    return os.path.join(base, "extraction_mac")


def chemin_config():
    return os.path.join(dossier_config(), NOM_FICHIER)


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
    """Retourne (profils, dernier_profil). Ne leve jamais : un fichier illisible
    est traite comme une configuration vide."""
    try:
        with open(chemin_config(), encoding="utf-8") as fichier:
            donnees = json.load(fichier)
    except (OSError, ValueError):
        return {}, ""
    if not isinstance(donnees, dict):
        return {}, ""
    bruts = donnees.get("profils") or donnees.get("profiles") or {}
    profils = {str(nom): normaliser_profil(valeur)
               for nom, valeur in bruts.items() if isinstance(nom, str)}
    dernier = donnees.get("dernier_profil") or ""
    return profils, (dernier if dernier in profils else "")


def enregistrer(profils, dernier_profil=""):
    """Ecrit la configuration de facon atomique. Leve OSError en cas d'echec."""
    dossier = dossier_config()
    os.makedirs(dossier, exist_ok=True)
    donnees = {
        "version": 1,
        "dernier_profil": dernier_profil,
        "profils": {nom: normaliser_profil(profil) for nom, profil in profils.items()},
    }
    descripteur, temporaire = tempfile.mkstemp(suffix=".tmp", dir=dossier)
    try:
        with os.fdopen(descripteur, "w", encoding="utf-8") as fichier:
            json.dump(donnees, fichier, indent=2, ensure_ascii=False)
        os.replace(temporaire, chemin_config())
    except BaseException:
        if os.path.exists(temporaire):
            os.remove(temporaire)
        raise
