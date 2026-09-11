"""Fenêtre d'ExtractionMAC : les profils, les deux onglets, la barre d'état.

Le programme fait deux métiers qui se suivent :

* **Relevé au banc** — capter l'adresse MAC de chaque appareil branché ;
* **Écriture dans Excel** — porter ces adresses dans la bonne cellule de la
  bonne fiche de test, qu'elles viennent du banc ou des fichiers Word.

Les réglages communs — le dossier des fiches, le fichier de la liste — sont
portés par des variables partagées entre les deux onglets : les saisir d'un côté
les renseigne de l'autre. Les profils sont gérés ici, une seule fois, et chaque
onglet se contente d'exposer ``lire_reglages`` / ``appliquer_reglages``.
"""

from __future__ import annotations

import os
import sys
import tkinter as tk
from tkinter import messagebox, ttk

from . import __version__, config
from .gui_banc import OngletBanc
from .gui_excel import OngletExcel
from .simulateur import SimulateurArp

NOM_DIAGNOSTIC = "ExtractionMAC-diagnostic.txt"


class Application(ttk.Frame):
    def __init__(self, racine, simulation=False):
        super().__init__(racine, padding=10)
        self.racine = racine
        self.grid(row=0, column=0, sticky="nsew")
        racine.columnconfigure(0, weight=1)
        racine.rowconfigure(0, weight=1)
        self.columnconfigure(0, weight=1)
        self.rowconfigure(2, weight=1)

        self.profils, dernier = config.charger()
        self._reglages = config.profil_vide()

        # partagées par les deux onglets : un même dossier, une même liste
        self.var_dossier = tk.StringVar()
        self.var_csv = tk.StringVar()
        self.var_profil = tk.StringVar()
        self.var_etat = tk.StringVar(value="Prêt.")

        self._construire(simulation)
        self._appliquer_profil(self.profils.get(dernier, config.profil_vide()))
        self.var_profil.set(dernier)
        self._rafraichir_profils()

        depart = "Configuration : %s" % config.chemin_config()
        if simulation:
            depart = "Mode simulation : aucun appareil réel n'est lu.   " + depart
        self.dire(depart)
        self.racine.protocol("WM_DELETE_WINDOW", self._fermer)

    # ------------------------------------------------------------------ vue
    def _construire(self, simulation):
        cadre_profil = ttk.Frame(self)
        cadre_profil.grid(row=0, column=0, sticky="ew")
        cadre_profil.columnconfigure(1, weight=1)
        ttk.Label(cadre_profil, text="Profil").grid(row=0, column=0, sticky="w")
        self.choix_profil = ttk.Combobox(cadre_profil, textvariable=self.var_profil, values=[])
        self.choix_profil.grid(row=0, column=1, sticky="ew", padx=(8, 0))
        self.choix_profil.bind("<<ComboboxSelected>>", self._charger_profil)
        ttk.Button(cadre_profil, text="Enregistrer", width=12,
                   command=self._enregistrer_profil).grid(row=0, column=2, padx=(6, 0))
        ttk.Button(cadre_profil, text="Supprimer", width=11,
                   command=self._supprimer_profil).grid(row=0, column=3, padx=(6, 0))

        ttk.Separator(self, orient="horizontal").grid(row=1, column=0, sticky="ew", pady=8)

        self.onglets = ttk.Notebook(self)
        self.onglets.grid(row=2, column=0, sticky="nsew")
        self.onglet_banc = OngletBanc(self.onglets, self,
                                      source=SimulateurArp() if simulation else None)
        self.onglet_excel = OngletExcel(self.onglets, self)
        self.onglets.add(self.onglet_banc, text="  Relevé au banc  ")
        self.onglets.add(self.onglet_excel, text="  Écriture dans Excel  ")

        ttk.Label(self, textvariable=self.var_etat, relief="sunken", anchor="w",
                  padding=(6, 3)).grid(row=3, column=0, sticky="ew", pady=(8, 0))

    # ------------------------------------------------- services aux onglets
    def dire(self, message):
        self.var_etat.set(message)

    def reglages(self):
        """Le profil courant complet : réglages des deux onglets et clés avancées."""
        profil = dict(self._reglages)
        profil["dossier"] = self.var_dossier.get().strip()
        profil["fichier_csv"] = self.var_csv.get().strip()
        profil.update(self.onglet_excel.lire_reglages())
        profil.update(self.onglet_banc.lire_reglages())
        return config.normaliser_profil(profil)

    def basculer_vers_ecriture(self):
        """Depuis l'onglet du banc : passer à l'écriture, source déjà choisie."""
        self.onglet_excel.choisir_source_releve()
        self.onglets.select(self.onglet_excel)
        self.dire("Source réglée sur la liste relevée : vérifier la feuille et la cellule, "
                  "puis « Analyser (sans écrire) ».")

    # --------------------------------------------------------------- profils
    def _appliquer_profil(self, profil):
        profil = config.normaliser_profil(profil)
        self._reglages = profil
        self.var_dossier.set(profil["dossier"])
        self.var_csv.set(profil["fichier_csv"])
        self.onglet_excel.appliquer_reglages(profil)
        self.onglet_banc.appliquer_reglages(profil)

    def _rafraichir_profils(self):
        self.choix_profil.configure(values=sorted(self.profils))

    def _charger_profil(self, _evenement=None):
        nom = self.var_profil.get().strip()
        if nom in self.profils:
            self._appliquer_profil(self.profils[nom])
            self.var_profil.set(nom)
            self.dire("Profil « %s » chargé." % nom)
            self.onglet_excel._lire_onglets(silencieux=True)

    def _enregistrer_profil(self):
        nom = self.var_profil.get().strip()
        if not nom:
            messagebox.showwarning("Profil", "Saisir un nom de profil avant d'enregistrer.",
                                   parent=self.racine)
            return
        self.profils[nom] = self.reglages()
        try:
            destination = config.enregistrer(self.profils, nom)
        except OSError as erreur:
            messagebox.showerror("Profil", "Enregistrement impossible :\n%s" % erreur,
                                 parent=self.racine)
            return
        self._rafraichir_profils()
        self.dire("Profil « %s » enregistré dans %s" % (nom, destination))

    def _supprimer_profil(self):
        nom = self.var_profil.get().strip()
        if nom not in self.profils:
            return
        if not messagebox.askyesno("Profil", "Supprimer le profil « %s » ?" % nom,
                                   parent=self.racine):
            return
        del self.profils[nom]
        try:
            config.enregistrer(self.profils, "")
        except OSError as erreur:
            messagebox.showerror("Profil", "Enregistrement impossible :\n%s" % erreur,
                                 parent=self.racine)
            return
        self.var_profil.set("")
        self._rafraichir_profils()
        self.dire("Profil « %s » supprimé." % nom)

    # ------------------------------------------------------------- fermeture
    def _fermer(self):
        for onglet in (self.onglet_excel, self.onglet_banc):
            if not onglet.fermer():
                return
        self.racine.destroy()


# --------------------------------------------------------------------------
# points d'entree
# --------------------------------------------------------------------------

def diagnostic():
    """Ecrit les emplacements resolus dans un fichier, a cote de la configuration.

    Sert au support (« ou sont passes mes profils ? ») et permet de verifier,
    sur l'executable reellement produit, que la configuration est bien ecrite
    a cote de lui.
    """
    lignes = [
        "Extraction MAC %s" % __version__,
        "fige par PyInstaller : %s" % bool(getattr(sys, "frozen", False)),
        "executable           : %s" % sys.executable,
        "dossier du programme : %s" % config.dossier_programme(),
        "dossier de repli     : %s" % config.dossier_repli(),
        "fichier de config    : %s" % config.chemin_config(),
    ]
    texte = "\n".join(lignes)
    cible = os.path.join(os.path.dirname(config.chemin_config()), NOM_DIAGNOSTIC)
    try:
        with open(cible, "w", encoding="utf-8") as fichier:
            fichier.write(texte + "\n")
    except OSError as erreur:
        texte += "\n(diagnostic non enregistre : %s)" % erreur
    if sys.stdout is not None:                     # absent en mode fenetre
        print(texte)
    return texte


def principal(argv=None):
    argv = sys.argv[1:] if argv is None else list(argv)
    options = {a.lstrip("-/").lower() for a in argv}
    if "diagnostic" in options:
        diagnostic()
        return 0
    lancer(simulation="simulation" in options)
    return 0


def lancer(simulation=False):
    racine = tk.Tk()
    titre = "Extraction MAC - relevé au banc et écriture dans Excel  v%s" % __version__
    if simulation:
        titre += "   [SIMULATION]"
    racine.title(titre)
    racine.minsize(880, 780)
    try:
        racine.call("ttk::style", "theme", "use", "vista")
    except tk.TclError:
        pass
    Application(racine, simulation=simulation)
    racine.mainloop()
