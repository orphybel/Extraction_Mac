"""Appairage des fichiers et report des adresses MAC dans les fiches Excel.

Deux sources d'adresses MAC aboutissent ici, et **elles partagent tout sauf la
premiere ligne** :

* les **fichiers Word** du dossier, appaires aux .xlsx par le numero de serie
  (``appairer`` / ``executer``) — c'est le fonctionnement historique ;
* la **liste relevee au banc de test** (``appairer_releves`` /
  ``executer_releves``).

Lecture de la cellule, cas « DEJA OK », protection contre l'ecrasement,
simulation et ecriture sont communs : c'est ``_executer_paires`` qui les porte,
et la source n'intervient que par la fonction ``obtenir_mac`` qu'on lui passe.
"""

from __future__ import annotations

import os
from dataclasses import dataclass, field

from . import arp, docmac, serial, xlsxcell

EXTENSIONS_WORD = (".doc", ".docx", ".docm", ".rtf")
EXTENSIONS_EXCEL = (".xlsx", ".xlsm")
EXTENSIONS_EXCEL_REFUSEES = (".xls", ".xlt")

ECRIT = "ecrit"
DEJA_OK = "deja_ok"
IGNORE = "ignore"
ERREUR = "erreur"
SIMULATION = "simulation"

_LIBELLES = {
    ECRIT: "ÉCRIT",
    DEJA_OK: "DÉJÀ OK",
    IGNORE: "IGNORÉ",
    ERREUR: "ERREUR",
    SIMULATION: "À ÉCRIRE",
}

# formes (singulier, pluriel) utilisees dans la ligne de bilan
_BILAN = {
    ECRIT: ("écrit", "écrits"),
    SIMULATION: ("à écrire", "à écrire"),
    DEJA_OK: ("déjà à jour", "déjà à jour"),
    IGNORE: ("ignoré", "ignorés"),
    ERREUR: ("erreur", "erreurs"),
}


@dataclass
class Options:
    dossier: str
    feuille: str
    cellule: str
    sous_dossiers: bool = False
    ecraser: bool = False
    simulation: bool = False
    separateur_mac: str = ":"
    mac_majuscules: bool = True
    groupes_numero: tuple = (6, 7, 5)


@dataclass
class Resultat:
    numero: str
    statut: str
    message: str
    word: str = ""
    excel: str = ""
    mac: str = ""

    @property
    def libelle(self):
        return _LIBELLES.get(self.statut, self.statut)


@dataclass
class Bilan:
    resultats: list = field(default_factory=list)

    def compter(self, statut):
        return sum(1 for r in self.resultats if r.statut == statut)

    def resume(self):
        parties = []
        for statut in (ECRIT, SIMULATION, DEJA_OK, IGNORE, ERREUR):
            nombre = self.compter(statut)
            if nombre:
                singulier, pluriel = _BILAN[statut]
                parties.append("%d %s" % (nombre, singulier if nombre == 1 else pluriel))
        return ", ".join(parties) if parties else "aucun fichier traité"


def _fichiers(dossier, sous_dossiers):
    if sous_dossiers:
        for racine, _, noms in os.walk(dossier):
            for nom in sorted(noms):
                yield os.path.join(racine, nom)
    else:
        for nom in sorted(os.listdir(dossier)):
            chemin = os.path.join(dossier, nom)
            if os.path.isfile(chemin):
                yield chemin


def _pertinent(nom):
    return not nom.startswith("~$") and not nom.startswith(".")


def numeros_des_fiches(dossier, sous_dossiers=False, groupes=(6, 7, 5)):
    """{numero: chemin} des fiches Excel presentes dans *dossier*.

    Sert au banc de test a controler un numero saisi : si aucune fiche ne le
    porte, c'est presque toujours une faute de frappe sur les derniers chiffres.
    """
    trouves = {}
    if not dossier or not os.path.isdir(dossier):
        return trouves
    for chemin in _fichiers(dossier, sous_dossiers):
        nom = os.path.basename(chemin)
        if not _pertinent(nom):
            continue
        if os.path.splitext(nom)[1].lower() not in EXTENSIONS_EXCEL:
            continue
        numero, _ = serial.key_from_filename(nom, groupes)
        if numero:
            trouves.setdefault(numero, chemin)
    return trouves


def _indexer(options):
    """Retourne (words, excels, anomalies) : {numero: [chemins]} pour le dossier."""
    groupes = tuple(options.groupes_numero)
    words, excels, anomalies = {}, {}, []

    for chemin in _fichiers(options.dossier, options.sous_dossiers):
        nom = os.path.basename(chemin)
        if not _pertinent(nom):
            continue
        extension = os.path.splitext(nom)[1].lower()
        if extension in EXTENSIONS_WORD:
            cible = words
        elif extension in EXTENSIONS_EXCEL:
            cible = excels
        elif extension in EXTENSIONS_EXCEL_REFUSEES:
            anomalies.append(Resultat("", ERREUR,
                                      "format .xls non géré ; enregistrer le fichier en .xlsx",
                                      excel=chemin))
            continue
        else:
            continue
        numero, erreur = serial.key_from_filename(nom, groupes)
        if erreur:
            anomalies.append(Resultat("", ERREUR, erreur,
                                      **({"word": chemin} if cible is words else {"excel": chemin})))
            continue
        cible.setdefault(numero, []).append(chemin)
    return words, excels, anomalies


def appairer(options):
    """Retourne (paires, anomalies) pour la source « fichiers Word ».

    ``paires`` : liste de (numero, chemin_word, chemin_excel).
    ``anomalies`` : liste de Resultat decrivant ce qui n'a pas pu etre appaire.
    """
    words, excels, anomalies = _indexer(options)

    paires = []
    for numero in sorted(set(words) | set(excels)):
        liste_word = words.get(numero, [])
        liste_excel = excels.get(numero, [])
        if not liste_excel:
            anomalies.append(Resultat(numero, IGNORE, "aucun fichier Excel portant ce numéro",
                                      word=liste_word[0]))
            continue
        if not liste_word:
            anomalies.append(Resultat(numero, IGNORE, "aucun fichier Word portant ce numéro",
                                      excel=liste_excel[0]))
            continue
        if len(liste_word) > 1:
            anomalies.append(Resultat(
                numero, ERREUR,
                "%d fichiers Word portent ce numéro : %s"
                % (len(liste_word), ", ".join(os.path.basename(c) for c in liste_word)),
                word=liste_word[0]))
            continue
        if len(liste_excel) > 1:
            anomalies.append(Resultat(
                numero, ERREUR,
                "%d fichiers Excel portent ce numéro : %s"
                % (len(liste_excel), ", ".join(os.path.basename(c) for c in liste_excel)),
                word=liste_word[0], excel=liste_excel[0]))
            continue
        paires.append((numero, liste_word[0], liste_excel[0]))
    return paires, anomalies


def appairer_releves(options, numeros):
    """Retourne (paires, anomalies) pour la source « liste relevée au banc ».

    ``numeros`` est la liste des numeros releves, dans l'ordre. Aucun fichier
    Word n'est attendu : la paire est (numero, "", chemin_excel).

    Difference assumee avec ``appairer`` : **une fiche Excel sans relevé n'est
    pas signalée**. Les deux sources n'ont pas la meme intention — le chemin
    Word traite tout un dossier, le chemin banc ecrit les quelques appareils
    qu'on vient de passer. Signaler les 200 fiches non concernees noierait le
    journal et ferait passer les vraies anomalies inapercues.
    """
    _, excels, anomalies = _indexer(options)

    paires, vus = [], set()
    for numero in numeros:
        if numero in vus:
            continue
        vus.add(numero)
        liste_excel = excels.get(numero, [])
        if not liste_excel:
            anomalies.append(Resultat(numero, IGNORE,
                                      "aucune fiche Excel portant ce numéro dans le dossier"))
            continue
        if len(liste_excel) > 1:
            anomalies.append(Resultat(
                numero, ERREUR,
                "%d fichiers Excel portent ce numéro : %s"
                % (len(liste_excel), ", ".join(os.path.basename(c) for c in liste_excel)),
                excel=liste_excel[0]))
            continue
        paires.append((numero, "", liste_excel[0]))
    return paires, anomalies


def _traiter(numero, chemin_word, chemin_excel, options, obtenir_mac):
    """Ecrit (ou simule) la MAC d'un appareil dans sa fiche.

    Seule la premiere ligne depend de la source : tout le reste est commun.
    """
    mac, erreur = obtenir_mac(numero, chemin_word)
    if erreur:
        return Resultat(numero, ERREUR, erreur, chemin_word, chemin_excel)

    cellule = options.cellule
    try:
        actuelle = xlsxcell.lire_cellule(chemin_excel, options.feuille, cellule)
    except xlsxcell.ErreurExcel as probleme:
        return Resultat(numero, ERREUR, str(probleme), chemin_word, chemin_excel, mac)
    except OSError as probleme:
        return Resultat(numero, ERREUR, "Excel illisible : %s" % probleme,
                        chemin_word, chemin_excel, mac)

    if actuelle:
        if str(actuelle).strip().casefold() == mac.casefold():
            return Resultat(numero, DEJA_OK, "%s déjà présent en %s" % (mac, cellule),
                            chemin_word, chemin_excel, mac)
        if not options.ecraser:
            return Resultat(
                numero, IGNORE,
                "%s contient déjà %r (cocher « Écraser » pour remplacer par %s)"
                % (cellule, actuelle, mac),
                chemin_word, chemin_excel, mac)

    if options.simulation:
        return Resultat(numero, SIMULATION, "%s serait écrit en %s!%s"
                        % (mac, options.feuille, cellule), chemin_word, chemin_excel, mac)

    try:
        xlsxcell.ecrire_cellule(chemin_excel, options.feuille, cellule, mac)
    except xlsxcell.ErreurExcel as probleme:
        return Resultat(numero, ERREUR, str(probleme), chemin_word, chemin_excel, mac)
    except PermissionError:
        return Resultat(numero, ERREUR,
                        "fichier verrouillé (probablement ouvert dans Excel) ; le fermer puis relancer",
                        chemin_word, chemin_excel, mac)
    except OSError as probleme:
        return Resultat(numero, ERREUR, "écriture impossible : %s" % probleme,
                        chemin_word, chemin_excel, mac)
    return Resultat(numero, ECRIT, "%s écrit en %s!%s" % (mac, options.feuille, cellule),
                    chemin_word, chemin_excel, mac)


def _valider(options):
    if not options.dossier or not os.path.isdir(options.dossier):
        raise ValueError("dossier introuvable : %s" % (options.dossier or "(vide)"))
    if not options.feuille.strip():
        raise ValueError("le nom de la feuille Excel est obligatoire")
    options.cellule = xlsxcell.normaliser_ref(options.cellule)


def _executer_paires(paires, anomalies, options, obtenir_mac, progression, interruption):
    """Boucle de traitement commune aux deux sources."""
    bilan = Bilan(list(anomalies))
    total = len(paires)
    for index, (numero, chemin_word, chemin_excel) in enumerate(paires, start=1):
        if interruption is not None and interruption():
            break
        resultat = _traiter(numero, chemin_word, chemin_excel, options, obtenir_mac)
        bilan.resultats.append(resultat)
        if progression is not None:
            progression(index, total, resultat)
    return bilan


def executer(options, progression=None, interruption=None):
    """Traite tout le dossier depuis les fichiers Word et retourne un Bilan.

    *progression* est appele avec (fait, total, Resultat) apres chaque fichier.
    *interruption* est appele sans argument ; s'il renvoie True, le traitement s'arrete.
    """
    _valider(options)
    paires, anomalies = appairer(options)

    def depuis_le_word(_numero, chemin_word):
        return docmac.extraire_mac(chemin_word, options.separateur_mac, options.mac_majuscules)

    return _executer_paires(paires, anomalies, options, depuis_le_word,
                            progression, interruption)


def executer_releves(options, releves, progression=None, interruption=None):
    """Ecrit les MAC relevees au banc dans les fiches Excel correspondantes.

    *releves* est une liste d'objets portant ``.numero`` et ``.mac``
    (voir ``releve.Releve``). Le Bilan et les statuts sont les memes que pour la
    source Word : le journal de l'interface n'a pas a savoir d'ou vient la MAC.
    """
    _valider(options)

    par_numero, conflits = {}, {}
    for element in releves:
        mac = arp.normaliser(element.mac, options.separateur_mac, options.mac_majuscules)
        ancienne = par_numero.get(element.numero)
        if ancienne is not None and ancienne != mac:
            conflits[element.numero] = (ancienne, mac)
        par_numero[element.numero] = mac

    paires, anomalies = appairer_releves(options, [r.numero for r in releves])
    paires = [p for p in paires if p[0] not in conflits]
    for numero, (premiere, seconde) in sorted(conflits.items()):
        anomalies.append(Resultat(
            numero, ERREUR,
            "deux adresses MAC différentes relevées pour ce numéro : %s et %s ; "
            "corriger la liste avant d'écrire" % (premiere, seconde)))

    def depuis_le_releve(numero, _chemin_word):
        mac = par_numero.get(numero)
        if not mac:
            return None, "aucune adresse MAC relevée pour ce numéro"
        exploitable, raison = arp.mac_exploitable(mac)
        if not exploitable:
            return None, raison
        return mac, None

    return _executer_paires(paires, anomalies, options, depuis_le_releve,
                            progression, interruption)
