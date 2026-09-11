"""Onglet « Écriture dans Excel » : porter les MAC dans les fiches de test.

C'est l'ecran historique d'ExtractionMAC, avec un seul ajout : le choix de la
provenance des adresses.

* **Fichiers Word du dossier** — le fonctionnement d'origine, inchange ;
* **Liste relevee au banc** — les adresses relevees dans l'autre onglet.

Le traitement tourne dans un fil separe : sur plusieurs centaines de fiches,
l'ecriture prend du temps et la fenetre doit rester vivante (barre de
progression, journal qui defile, bouton « Arreter »).
"""

from __future__ import annotations

import os
import queue
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk

from . import releve, runner, xlsxcell

COULEURS = {
    runner.ECRIT: "#0a7a28",
    runner.SIMULATION: "#1a5fb4",
    runner.DEJA_OK: "#5c5c5c",
    runner.IGNORE: "#a35c00",
    runner.ERREUR: "#c01c28",
}

SOURCE_WORD = "word"
SOURCE_RELEVE = "releve"


class OngletExcel(ttk.Frame):
    def __init__(self, parent, application):
        super().__init__(parent, padding=10)
        self.app = application
        self.columnconfigure(1, weight=1)

        self.messages = queue.Queue()
        self.travail = None
        self.arret = threading.Event()

        self.var_dossier = application.var_dossier          # partagé avec l'onglet banc
        self.var_csv = application.var_csv                  # partagé avec l'onglet banc
        self.var_feuille = tk.StringVar()
        self.var_cellule = tk.StringVar()
        self.var_sous_dossiers = tk.BooleanVar()
        self.var_ecraser = tk.BooleanVar()
        self.var_source = tk.StringVar(value=SOURCE_WORD)

        self._construire()
        self.after(100, self._vider_file)

    # ------------------------------------------------------------------ vue
    def _construire(self):
        ligne = 0

        ttk.Label(self, text="Dossier des fichiers").grid(row=ligne, column=0, sticky="w", pady=3)
        ttk.Entry(self, textvariable=self.var_dossier).grid(row=ligne, column=1, sticky="ew",
                                                            pady=3, padx=(8, 0))
        ttk.Button(self, text="Parcourir…", width=14,
                   command=self._choisir_dossier).grid(row=ligne, column=2, padx=(6, 0))
        ligne += 1

        ttk.Label(self, text="Feuille Excel").grid(row=ligne, column=0, sticky="w", pady=3)
        self.choix_feuille = ttk.Combobox(self, textvariable=self.var_feuille, values=[])
        self.choix_feuille.grid(row=ligne, column=1, sticky="ew", pady=3, padx=(8, 0))
        ttk.Button(self, text="Lire les onglets", width=14,
                   command=self._lire_onglets).grid(row=ligne, column=2, padx=(6, 0))
        ligne += 1

        ttk.Label(self, text="Cellule").grid(row=ligne, column=0, sticky="w", pady=3)
        cadre = ttk.Frame(self)
        cadre.grid(row=ligne, column=1, sticky="ew", pady=3, padx=(8, 0))
        cadre.columnconfigure(1, weight=1)
        ttk.Entry(cadre, textvariable=self.var_cellule, width=10).grid(row=0, column=0)
        ttk.Label(cadre, text="  (exemple : F27)",
                  foreground="#5c5c5c").grid(row=0, column=1, sticky="w")
        ttk.Button(self, text="Trouver @MAC", width=14,
                   command=self._trouver_libelles).grid(row=ligne, column=2, padx=(6, 0))
        ligne += 1

        # --- provenance des adresses ---------------------------------------
        cadre_source = ttk.LabelFrame(self, text="Source des adresses MAC", padding=8)
        cadre_source.grid(row=ligne, column=0, columnspan=3, sticky="ew", pady=(10, 4))
        cadre_source.columnconfigure(1, weight=1)
        ttk.Radiobutton(cadre_source, text="Fichiers Word du dossier",
                        variable=self.var_source, value=SOURCE_WORD,
                        command=self._peindre_source).grid(row=0, column=0, sticky="w")
        ttk.Radiobutton(cadre_source, text="Liste relevée au banc",
                        variable=self.var_source, value=SOURCE_RELEVE,
                        command=self._peindre_source).grid(row=1, column=0, sticky="w", pady=(2, 0))
        self.cadre_csv = ttk.Frame(cadre_source)
        self.cadre_csv.grid(row=1, column=1, sticky="ew", padx=(16, 0))
        self.cadre_csv.columnconfigure(0, weight=1)
        ttk.Entry(self.cadre_csv, textvariable=self.var_csv).grid(row=0, column=0, sticky="ew")
        ttk.Button(self.cadre_csv, text="Parcourir…", width=12,
                   command=self._choisir_csv).grid(row=0, column=1, padx=(6, 0))
        ligne += 1

        cadre_options = ttk.Frame(self)
        cadre_options.grid(row=ligne, column=0, columnspan=3, sticky="w", pady=(4, 0))
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
        self.journal = tk.Text(cadre_journal, height=14, wrap="word", state="disabled")
        self.journal.grid(row=0, column=0, sticky="nsew")
        defilement = ttk.Scrollbar(cadre_journal, orient="vertical", command=self.journal.yview)
        defilement.grid(row=0, column=1, sticky="ns")
        self.journal.configure(yscrollcommand=defilement.set)
        for statut, couleur in COULEURS.items():
            self.journal.tag_configure(statut, foreground=couleur)
        self.journal.tag_configure("titre", font=("TkDefaultFont", 9, "bold"))

        self._peindre_source()

    def _peindre_source(self):
        actif = self.var_source.get() == SOURCE_RELEVE
        for enfant in self.cadre_csv.winfo_children():
            enfant.configure(state="normal" if actif else "disabled")

    # -------------------------------------------------------------- profils
    def lire_reglages(self):
        return {
            "feuille": self.var_feuille.get().strip(),
            "cellule": self.var_cellule.get().strip(),
            "sous_dossiers": self.var_sous_dossiers.get(),
            "ecraser": self.var_ecraser.get(),
            "source_mac": self.var_source.get(),
        }

    def appliquer_reglages(self, profil):
        self.var_feuille.set(profil["feuille"])
        self.var_cellule.set(profil["cellule"])
        self.var_sous_dossiers.set(profil["sous_dossiers"])
        self.var_ecraser.set(profil["ecraser"])
        self.var_source.set(profil["source_mac"])
        self._peindre_source()

    def choisir_source_releve(self):
        """Appelé depuis l'onglet du banc : écrire ce qui vient d'être relevé."""
        self.var_source.set(SOURCE_RELEVE)
        self._peindre_source()

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

    # ----------------------------------------------------------- assistance
    def _choisir_dossier(self):
        depart = self.var_dossier.get() if os.path.isdir(self.var_dossier.get()) else None
        dossier = filedialog.askdirectory(title="Dossier contenant les fiches de test",
                                          initialdir=depart, parent=self)
        if dossier:
            self.var_dossier.set(os.path.normpath(dossier))
            self._lire_onglets(silencieux=True)

    def _choisir_csv(self):
        choix = filedialog.askopenfilename(
            title="Liste relevée au banc", defaultextension=".csv",
            filetypes=[("Fichier CSV", "*.csv"), ("Tous les fichiers", "*.*")], parent=self)
        if choix:
            self.var_csv.set(choix)

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
                                    parent=self)
            return
        try:
            onglets = xlsxcell.lister_feuilles(classeur)
        except (xlsxcell.ErreurExcel, OSError) as erreur:
            if not silencieux:
                messagebox.showerror("Onglets", str(erreur), parent=self)
            return
        self.choix_feuille.configure(values=onglets)
        if self.var_feuille.get().strip() not in onglets:
            self.app.dire("Onglets lus dans %s : vérifier la feuille choisie."
                          % os.path.basename(classeur))
        else:
            self.app.dire("Onglets lus dans %s." % os.path.basename(classeur))

    def _trouver_libelles(self):
        classeur = self._premier_classeur()
        if not classeur:
            messagebox.showinfo("Recherche", "Aucun fichier .xlsx trouvé dans ce dossier.",
                                parent=self)
            return
        try:
            trouves = xlsxcell.trouver_libelles(classeur, "MAC")
        except (xlsxcell.ErreurExcel, OSError) as erreur:
            messagebox.showerror("Recherche", str(erreur), parent=self)
            return
        self._ecrire("Libellés contenant « MAC » dans %s :" % os.path.basename(classeur), "titre")
        if not trouves:
            self._ecrire("  aucun libellé trouvé.", runner.IGNORE)
            return
        for feuille, cellule, texte in trouves:
            self._ecrire("  %s!%s : %s" % (feuille, cellule, texte))
        self._ecrire("  La MAC se met en général sur la même ligne, dans la colonne « N° Série ».",
                     runner.DEJA_OK)

    # ------------------------------------------------------------ execution
    def _options(self, simulation):
        reglages = self.app.reglages()
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
        return options

    def _releves(self):
        """Liste relevée au banc, relue depuis le CSV. Lève ValueError si absente."""
        chemin = self.var_csv.get().strip()
        if not chemin:
            raise ValueError("aucun fichier de relevés choisi")
        if not os.path.isfile(chemin):
            raise ValueError("fichier de relevés introuvable : %s" % chemin)
        releves, message = releve.charger(chemin)
        if not releves:
            raise ValueError(message or "la liste relevée est vide")
        return releves

    def _lancer(self, simulation):
        if self.travail and self.travail.is_alive():
            return
        depuis_releve = self.var_source.get() == SOURCE_RELEVE
        try:
            options = self._options(simulation)
            releves = self._releves() if depuis_releve else None
        except (ValueError, xlsxcell.ErreurExcel) as erreur:
            messagebox.showerror("Paramètres", str(erreur), parent=self)
            return

        if not simulation:
            if depuis_releve:
                paires, _ = runner.appairer_releves(options, [r.numero for r in releves])
                rien = ("Aucun relevé de la liste ne correspond à une fiche Excel "
                        "de ce dossier.")
            else:
                paires, _ = runner.appairer(options)
                rien = "Aucune paire .doc / .xlsx à traiter dans ce dossier."
            if not paires:
                messagebox.showinfo("Écriture", rien, parent=self)
                return
            if not messagebox.askyesno(
                    "Confirmation",
                    "%d fichier(s) Excel vont être modifiés sur place,\n"
                    "en %s!%s.\n\nContinuer ?" % (len(paires), options.feuille, options.cellule),
                    parent=self):
                return

        self._effacer()
        self._ecrire("%s - %s!%s%s"
                     % ("Analyse (aucune écriture)" if simulation else "Écriture",
                        options.feuille, options.cellule,
                        " - sous-dossiers inclus" if options.sous_dossiers else ""), "titre")
        self._ecrire("Source : %s" % ("liste relevée au banc (%d relevé(s))" % len(releves)
                                      if depuis_releve else "fichiers Word du dossier"))
        self._ecrire(options.dossier)
        self._ecrire("")
        self.arret.clear()
        self.progression.configure(value=0, maximum=100)
        self._basculer(False)
        self.travail = threading.Thread(target=self._executer, args=(options, releves), daemon=True)
        self.travail.start()

    def _executer(self, options, releves):
        def progression(fait, total, resultat):
            self.messages.put(("avance", (fait, total, resultat)))
        try:
            if releves is None:
                bilan = runner.executer(options, progression, self.arret.is_set)
            else:
                bilan = runner.executer_releves(options, releves, progression, self.arret.is_set)
        except Exception as erreur:                   # noqa: BLE001 - remonte a l'interface
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
                    messagebox.showerror("Erreur", charge, parent=self)
                    self.app.dire("Échec : %s" % charge)
        except queue.Empty:
            pass
        self.after(100, self._vider_file)

    def _afficher(self, resultat):
        cible = os.path.basename(resultat.excel or resultat.word or "")
        self._ecrire("[%-9s] %-22s %s" % (resultat.libelle, resultat.numero or "-", cible),
                     resultat.statut)
        self._ecrire("             %s" % resultat.message, resultat.statut)

    def _terminer(self, bilan):
        anomalies = [r for r in bilan.resultats
                     if r.statut in (runner.ERREUR, runner.IGNORE) and not r.mac]
        if anomalies:
            self._ecrire("")
            self._ecrire("Fichiers non traités :", "titre")
            for resultat in anomalies:
                self._afficher(resultat)
        self._ecrire("")
        self._ecrire("Bilan : %s" % bilan.resume(), "titre")
        self.app.dire("Terminé - %s" % bilan.resume())
        self.progression.configure(value=self.progression["maximum"])
        self._basculer(True)

    def _basculer(self, actif):
        etat = "normal" if actif else "disabled"
        self.bouton_analyse.configure(state=etat)
        self.bouton_ecriture.configure(state=etat)
        self.bouton_arret.configure(state="disabled" if actif else "normal")

    def fermer(self):
        """Retourne False pour annuler la fermeture de la fenêtre."""
        if self.travail and self.travail.is_alive():
            if not messagebox.askyesno("Quitter", "Un traitement est en cours. Quitter quand même ?",
                                       parent=self):
                return False
            self.arret.set()
        return True
