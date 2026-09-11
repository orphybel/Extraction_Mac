"""La liste des appareils releves, et son fichier CSV.

Le CSV est la **liste de reference** : c'est lui que l'on relit, que l'on
archive et que l'on ouvre dans Excel. Il est ecrit avec le point-virgule et en
UTF-8 avec BOM pour qu'Excel en francais l'ouvre correctement d'un double-clic,
sans passer par l'assistant d'importation.

Il est relu au demarrage : la liste survit a une fermeture du programme, et les
doublons sont detectes d'une seance sur l'autre.
"""

from __future__ import annotations

import csv
import datetime
import os
import tempfile
from dataclasses import dataclass, field

COLONNES = ("numero", "mac", "ip", "horodatage")
NOM_DEFAUT = "MAC-releves.csv"


def maintenant():
    return datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")


@dataclass
class Releve:
    numero: str
    mac: str
    ip: str = ""
    horodatage: str = field(default_factory=maintenant)

    def en_ligne(self):
        return {"numero": self.numero, "mac": self.mac,
                "ip": self.ip, "horodatage": self.horodatage}


def controler(releves, numero, mac):
    """Anomalies a signaler avant d'ajouter (numero, mac) a la liste.

    Ne bloque pas : l'operateur reste juge. Mais rien n'est ajoute en silence.
    """
    anomalies = []
    for existant in releves:
        if existant.numero == numero:
            anomalies.append("le numéro %s est déjà dans la liste (MAC %s)"
                             % (numero, existant.mac))
        if existant.mac == mac and existant.numero != numero:
            anomalies.append("cette MAC est déjà relevée sous le numéro %s"
                             % existant.numero)
    return anomalies


def _dialecte(premiere_ligne):
    """Point-virgule par defaut ; la virgule est acceptee en relecture."""
    if ";" in premiere_ligne or "," not in premiere_ligne:
        return ";"
    return ","


def charger(chemin):
    """Retourne (releves, message). Un fichier absent n'est pas une erreur."""
    if not chemin or not os.path.isfile(chemin):
        return [], ""
    try:
        with open(chemin, encoding="utf-8-sig", newline="") as fichier:
            texte = fichier.read()
    except OSError as erreur:
        return [], "liste illisible : %s" % erreur
    if not texte.strip():
        return [], ""

    lignes = texte.splitlines()
    lecteur = csv.DictReader(lignes, delimiter=_dialecte(lignes[0]))
    if not lecteur.fieldnames or "numero" not in lecteur.fieldnames:
        return [], "en-tête inattendu dans %s (colonnes attendues : %s)" % (
            os.path.basename(chemin), ";".join(COLONNES))

    releves, ignorees = [], 0
    for ligne in lecteur:
        numero = (ligne.get("numero") or "").strip()
        mac = (ligne.get("mac") or "").strip()
        if not numero or not mac:
            ignorees += 1
            continue
        releves.append(Releve(numero, mac,
                              (ligne.get("ip") or "").strip(),
                              (ligne.get("horodatage") or "").strip()))
    message = "%d relevé(s) repris de %s" % (len(releves), os.path.basename(chemin))
    if ignorees:
        message += " (%d ligne(s) incomplète(s) ignorée(s))" % ignorees
    return releves, message


def enregistrer(chemin, releves):
    """Ecrit la liste de facon atomique et retourne le chemin utilise.

    L'ecriture passe par un fichier temporaire puis ``os.replace`` : une coupure
    au mauvais moment ne peut pas laisser une liste tronquee a la place de la
    liste complete.
    """
    dossier = os.path.dirname(os.path.abspath(chemin)) or "."
    os.makedirs(dossier, exist_ok=True)
    descripteur, temporaire = tempfile.mkstemp(suffix=".tmp", dir=dossier)
    try:
        with os.fdopen(descripteur, "w", encoding="utf-8-sig", newline="") as fichier:
            redacteur = csv.DictWriter(fichier, fieldnames=list(COLONNES), delimiter=";")
            redacteur.writeheader()
            for releve in releves:
                redacteur.writerow(releve.en_ligne())
        os.replace(temporaire, chemin)
    except BaseException:
        if os.path.exists(temporaire):
            os.remove(temporaire)
        raise
    return chemin
