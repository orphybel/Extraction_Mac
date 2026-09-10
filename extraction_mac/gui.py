"""Fenetre Tkinter : choix du dossier, de l'onglet et de la cellule, puis report des MAC."""

from __future__ import annotations

import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import __version__, config, runner, xlsxcell

COULEURS = {
    runner.ECRIT: "#0a7a28",
    runner.SIMULATION: "#1a5fb4",
    runner.DEJA_OK: "#5c5c5c",
    runner.IGNORE: "#a35c00",
    runner.ERREUR: "#c01c28",
}


class Application(ttk.Frame):
    def __init__(self, racine):
        super().__init__(racine, padding=12)
        self.racine = racine
        self.grid(row=0, column=0, sticky="nsew")
        racine.columnconfigure(0, weight=1)
        racine.rowconfigure(0, weight=1)
        self.columnconfigure(1, weight=1)

        self.profils, dernier = config.charger()
        self.messages = queue.Queue()
        self.travail = None
        self.arret = threading.Event()

        self.var_profil = tk.StringVar()
        self.var_dossier = tk.StringVar()
        self.var_feuille = tk.StringVar()
        self.var_cellule = tk.StringVar()
        self.var_sous_dossiers = tk.BooleanVar()
        self.var_ecraser = tk.BooleanVar()
        self.var_etat = tk.StringVar(value="Prêt.")

        self._construire()
        self._appliquer_profil(self.profils.get(dernier, config.profil_vide()))
        self.var_profil.set(dernier)
        self._rafraichir_profils()
        self.var_etat.set("Configuration : %s" % config.chemin_config())
        self.racine.after(100, self._vider_file)
        self.racine.protocol("WM_DELETE_WINDOW", self._fermer)

    # ------------------------------------------------------------------ vue
    def _construire(self):
        ligne = 0

        ttk.Label(self, text="Profil").grid(row=ligne, column=0, sticky="w", pady=3)
        cadre_profil = ttk.Frame(self)
        cadre_profil.grid(row=ligne, column=1, columnspan=2, sticky="ew", pady=3, padx=(8, 0))
        cadre_profil.columnconfigure(0, weight=1)
        self.choix_profil = ttk.Combobox(cadre_profil, textvariable=self.var_profil, values=[])
        self.choix_profil.grid(row=0, column=0, sticky="ew")
        self.choix_profil.bind("<<ComboboxSelected>>", self._charger_profil)
        ttk.Button(cadre_profil, text="Enregistrer", width=12,
                   command=self._enregistrer_profil).grid(row=0, column=1, padx=(6, 0))
        ttk.Button(cadre_profil, text="Supprimer", width=11,
                   command=self._supprimer_profil).grid(row=0, column=2, padx=(6, 0))
        ligne += 1

        ttk.Separator(self, orient="horizontal").grid(row=ligne, column=0, columnspan=3,
                                                      sticky="ew", pady=8)
        ligne += 1

        ttk.Label(self, text="Dossier des fichiers").grid(row=ligne, column=0, sticky="w", pady=3)
        ttk.Entry(self, textvariable=self.var_dossier).grid(row=ligne, column=1, sticky="ew",
                                                            pady=3, padx=(8, 0))
        ttk.Button(self, text="Parcourir...", width=13,
                   command=self._choisir_dossier).grid(row=ligne, column=2, padx=(6, 0))
        ligne += 1

        ttk.Label(self, text="Feuille Excel").grid(row=ligne, column=0, sticky="w", pady=3)
        self.choix_feuille = ttk.Combobox(self, textvariable=self.var_feuille, values=[])
        self.choix_feuille.grid(row=ligne, column=1, sticky="ew", pady=3, padx=(8, 0))
        ttk.Button(self, text="Lire les onglets", width=13,
                   command=self._lire_onglets).grid(row=ligne, column=2, padx=(6, 0))
        ligne += 1

        ttk.Label(self, text="Cellule").grid(row=ligne, column=0, sticky="w", pady=3)
        cadre_cellule = ttk.Frame(self)
        cadre_cellule.grid(row=ligne, column=1, sticky="ew", pady=3, padx=(8, 0))
        cadre_cellule.columnconfigure(1, weight=1)
        ttk.Entry(cadre_cellule, textvariable=self.var_cellule, width=10).grid(row=0, column=0)
        ttk.Label(cadre_cellule, text="  (exemple : F27)",
                  foreground="#5c5c5c").grid(row=0, column=1, sticky="w")
        ttk.Button(self, text="Trouver @MAC", width=13,
                   command=self._trouver_libelles).grid(row=ligne, column=2, padx=(6, 0))
        ligne += 1

        cadre_options = ttk.Frame(self)
        cadre_options.grid(row=ligne, column=1, columnspan=2, sticky="w", pady=(6, 0), padx=(8, 0))
        ttk.Checkbutton(cadre_options, text="Inclure les sous-dossiers",
                        variable=self.var_sous_dossiers).grid(row=0, column=0, sticky="w")
        ttk.Checkbutton(cadre_options, text="Écraser une valeur déjà présente",
                        variable=self.var_ecraser).grid(row=0, column=1, sticky="w", padx=(16, 0))
        ligne += 1

        cadre_actions = ttk.Frame(self)
        cadre_actions.grid(row=ligne, column=0, columnspan=3, sticky="ew", pady=(12, 6))
        cadre_actions.columnconfigure(2, weight=1)
        self.bouton_analyse = ttk.Button(cadre_actions, text="Analyser (sans écrire)",
                                         command=lambda: self._lancer(True))
        self.bouton_analyse.grid(row=0, column=0)
        self.bouton_ecriture = ttk.Button(cadre_actions, text="Écrire dans Excel",
                                          command=lambda: self._lancer(False))
        self.bouton_ecriture.grid(row=0, column=1, padx=(8, 0))
        self.bouton_arret = ttk.Button(cadre_actions, text="Arrêter",
                                       command=self.arret.set, state="disabled")
        self.bouton_arret.grid(row=0, column=2, sticky="w", padx=(8, 0))
        ttk.Button(cadre_actions, text="Effacer le journal",
                   command=self._effacer).grid(row=0, column=3, sticky="e")
        ligne += 1

        self.progression = ttk.Progressbar(self, mode="determinate")
        self.progression.grid(row=ligne, column=0, columnspan=3, sticky="ew")
        ligne += 1

        cadre_journal = ttk.Frame(self)
        cadre_journal.grid(row=ligne, column=0, columnspan=3, sticky="nsew", pady=(8, 0))
        cadre_journal.columnconfigure(0, weight=1)
        cadre_journal.rowconfigure(0, weight=1)
        self.rowconfigure(ligne, weight=1)
        self.journal = tk.Text(cadre_journal, height=16, wrap="word", state="disabled")
        self.journal.grid(row=0, column=0, sticky="nsew")
        defilement = ttk.Scrollbar(cadre_journal, orient="vertical", command=self.journal.yview)
        defilement.grid(row=0, column=1, sticky="ns")
        self.journal.configure(yscrollcommand=defilement.set)
        for statut, couleur in COULEURS.items():
            self.journal.tag_configure(statut, foreground=couleur)
        self.journal.tag_configure("titre", font=("TkDefaultFont", 9, "bold"))
        ligne += 1

        ttk.Label(self, textvariable=self.var_etat, foreground="#5c5c5c").grid(
            row=ligne, column=0, columnspan=3, sticky="w", pady=(6, 0))

    # -------------------------------------------------------------- journal
    def _ecrire(self, texte, tag=None):
        self.journal.configure(state="normal")
        self.journal.insert("end", texte + "\n", tag or ())
        self.journal.see("end")
        self.journal.configure(state="disabled")

    def _effacer(self):
        self.journal.configure(state="normal")
        self.journal.delete("1.0", "end")
        self.journal.configure(state="disabled")

    # --------------------------------------------------------------- profils
    def _profil_courant(self):
        return {
            "dossier": self.var_dossier.get().strip(),
            "feuille": self.var_feuille.get().strip(),
            "cellule": self.var_cellule.get().strip(),
            "sous_dossiers": self.var_sous_dossiers.get(),
            "ecraser": self.var_ecraser.get(),
        }

    def _appliquer_profil(self, profil):
        profil = config.normaliser_profil(profil)
        self.var_dossier.set(profil["dossier"])
        self.var_feuille.set(profil["feuille"])
        self.var_cellule.set(profil["cellule"])
        self.var_sous_dossiers.set(profil["sous_dossiers"])
        self.var_ecraser.set(profil["ecraser"])
        self._reglages = profil

    def _rafraichir_profils(self):
        self.choix_profil.configure(values=sorted(self.profils))

    def _charger_profil(self, _evenement=None):
        nom = self.var_profil.get().strip()
        if nom in self.profils:
            self._appliquer_profil(self.profils[nom])
            self.var_profil.set(nom)
            self.var_etat.set("Profil « %s » chargé." % nom)
            self._lire_onglets(silencieux=True)

    def _enregistrer_profil(self):
        nom = self.var_profil.get().strip()
        if not nom:
            messagebox.showwarning("Profil", "Saisir un nom de profil avant d'enregistrer.",
                                   parent=self.racine)
            return
        profil = dict(getattr(self, "_reglages", config.profil_vide()))
        profil.update(self._profil_courant())
        self.profils[nom] = config.normaliser_profil(profil)
        try:
            destination = config.enregistrer(self.profils, nom)
        except OSError as erreur:
            messagebox.showerror("Profil", "Enregistrement impossible :\n%s" % erreur,
                                 parent=self.racine)
            return
        self._rafraichir_profils()
        self.var_etat.set("Profil « %s » enregistré dans %s" % (nom, destination))

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
        self.var_etat.set("Profil « %s » supprimé." % nom)

    # ------------------------------------------------------------- assistance
    def _choisir_dossier(self):
        depart = self.var_dossier.get() if os.path.isdir(self.var_dossier.get()) else None
        dossier = filedialog.askdirectory(title="Dossier contenant les .doc et les .xlsx",
                                          initialdir=depart, parent=self.racine)
        if dossier:
            self.var_dossier.set(os.path.normpath(dossier))
            self._lire_onglets(silencieux=True)

    def _premier_classeur(self):
        dossier = self.var_dossier.get().strip()
        if not os.path.isdir(dossier):
            return None
        for chemin in runner._fichiers(dossier, self.var_sous_dossiers.get()):
            nom = os.path.basename(chemin)
            if runner._pertinent(nom) and nom.lower().endswith(runner.EXTENSIONS_EXCEL):
                return chemin
        return None

    def _lire_onglets(self, silencieux=False):
        classeur = self._premier_classeur()
        if not classeur:
            if not silencieux:
                messagebox.showinfo("Onglets", "Aucun fichier .xlsx trouvé dans ce dossier.",
                                    parent=self.racine)
            return
        try:
            onglets = xlsxcell.lister_feuilles(classeur)
        except (xlsxcell.ErreurExcel, OSError) as erreur:
            if not silencieux:
                messagebox.showerror("Onglets", str(erreur), parent=self.racine)
            return
        self.choix_feuille.configure(values=onglets)
        if self.var_feuille.get().strip() not in onglets:
            self.var_etat.set("Onglets lus dans %s : vérifier la feuille choisie."
                              % os.path.basename(classeur))
        else:
            self.var_etat.set("Onglets lus dans %s." % os.path.basename(classeur))

    def _trouver_libelles(self):
        classeur = self._premier_classeur()
        if not classeur:
            messagebox.showinfo("Recherche", "Aucun fichier .xlsx trouvé dans ce dossier.",
                                parent=self.racine)
            return
        try:
            trouves = xlsxcell.trouver_libelles(classeur, "MAC")
        except (xlsxcell.ErreurExcel, OSError) as erreur:
            messagebox.showerror("Recherche", str(erreur), parent=self.racine)
            return
        self._ecrire("Libellés contenant « MAC » dans %s :" % os.path.basename(classeur), "titre")
        if not trouves:
            self._ecrire("  aucun libellé trouvé.", runner.IGNORE)
            return
        for feuille, cellule, texte in trouves:
            self._ecrire("  %s!%s : %s" % (feuille, cellule, texte))
        self._ecrire("  La MAC se met en général sur la même ligne, dans la colonne « N° Série ».",
                     runner.DEJA_OK)

    # -------------------------------------------------------------- execution
    def _lancer(self, simulation):
        if self.travail and self.travail.is_alive():
            return
        reglages = dict(getattr(self, "_reglages", config.profil_vide()))
        reglages.update(self._profil_courant())
        try:
            options = runner.Options(
                dossier=reglages["dossier"],
                feuille=reglages["feuille"],
                cellule=reglages["cellule"],
                sous_dossiers=reglages["sous_dossiers"],
                ecraser=reglages["ecraser"],
                simulation=simulation,
                separateur_mac=reglages["separateur_mac"],
                mac_majuscules=reglages["mac_majuscules"],
                groupes_numero=tuple(reglages["groupes_numero"]),
            )
            options.cellule = xlsxcell.normaliser_ref(options.cellule)
            if not options.dossier or not os.path.isdir(options.dossier):
                raise ValueError("dossier introuvable : %s" % (options.dossier or "(vide)"))
            if not options.feuille.strip():
                raise ValueError("le nom de la feuille Excel est obligatoire")
        except (ValueError, xlsxcell.ErreurExcel) as erreur:
            messagebox.showerror("Paramètres", str(erreur), parent=self.racine)
            return

        if not simulation:
            paires, _ = runner.appairer(options)
            if not paires:
                messagebox.showinfo("Ecriture", "Aucune paire .doc / .xlsx à traiter dans ce dossier.",
                                    parent=self.racine)
                return
            if not messagebox.askyesno(
                    "Confirmation",
                    "%d fichier(s) Excel vont être modifiés sur place,\n"
                    "en %s!%s.\n\nContinuer ?" % (len(paires), options.feuille, options.cellule),
                    parent=self.racine):
                return

        self._effacer()
        self._ecrire("%s - %s!%s%s"
                     % ("Analyse (aucune écriture)" if simulation else "Écriture",
                        options.feuille, options.cellule,
                        " - sous-dossiers inclus" if options.sous_dossiers else ""), "titre")
        self._ecrire(options.dossier)
        self._ecrire("")
        self.arret.clear()
        self.progression.configure(value=0, maximum=100)
        self._basculer(False)
        self.travail = threading.Thread(target=self._executer, args=(options,), daemon=True)
        self.travail.start()

    def _executer(self, options):
        def progression(fait, total, resultat):
            self.messages.put(("avance", (fait, total, resultat)))
        try:
            bilan = runner.executer(options, progression, self.arret.is_set)
        except Exception as erreur:                       # noqa: BLE001 - remonte a l'interface
            self.messages.put(("echec", str(erreur)))
            return
        self.messages.put(("fin", bilan))

    def _vider_file(self):
        try:
            while True:
                genre, charge = self.messages.get_nowait()
                if genre == "avance":
                    fait, total, resultat = charge
                    self.progression.configure(maximum=max(total, 1), value=fait)
                    self._afficher(resultat)
                elif genre == "fin":
                    self._terminer(charge)
                elif genre == "echec":
                    self._basculer(True)
                    messagebox.showerror("Erreur", charge, parent=self.racine)
                    self.var_etat.set("Échec : %s" % charge)
        except queue.Empty:
            pass
        self.racine.after(100, self._vider_file)

    def _afficher(self, resultat):
        cible = os.path.basename(resultat.excel or resultat.word or "")
        self._ecrire("[%-9s] %-22s %s" % (resultat.libelle, resultat.numero or "-", cible),
                     resultat.statut)
        self._ecrire("             %s" % resultat.message, resultat.statut)

    def _terminer(self, bilan):
        anomalies = [r for r in bilan.resultats if r.statut in (runner.ERREUR, runner.IGNORE)
                     and not r.mac]
        if anomalies:
            self._ecrire("")
            self._ecrire("Fichiers non traités :", "titre")
            for resultat in anomalies:
                self._afficher(resultat)
        self._ecrire("")
        self._ecrire("Bilan : %s" % bilan.resume(), "titre")
        self.var_etat.set("Terminé - %s" % bilan.resume())
        self.progression.configure(value=self.progression["maximum"])
        self._basculer(True)

    def _basculer(self, actif):
        etat = "normal" if actif else "disabled"
        self.bouton_analyse.configure(state=etat)
        self.bouton_ecriture.configure(state=etat)
        self.bouton_arret.configure(state="disabled" if actif else "normal")

    def _fermer(self):
        if self.travail and self.travail.is_alive():
            if not messagebox.askyesno("Quitter", "Un traitement est en cours. Quitter quand même ?",
                                       parent=self.racine):
                return
            self.arret.set()
        self.racine.destroy()


NOM_DIAGNOSTIC = "ExtractionMAC-diagnostic.txt"


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
    if any(argument.lstrip("-/").lower() == "diagnostic" for argument in argv):
        diagnostic()
        return 0
    lancer()
    return 0


def lancer():
    racine = tk.Tk()
    racine.title("Extraction MAC - Word vers Excel  v%s" % __version__)
    racine.minsize(760, 620)
    try:
        racine.call("ttk::style", "theme", "use", "vista")
    except tk.TclError:
        pass
    Application(racine)
    racine.mainloop()
