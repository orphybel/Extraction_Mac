"""Onglet « Relevé au banc » : capter l'adresse MAC de chaque appareil branché.

Le geste de l'operateur doit tenir en une touche : l'appareil est detecte, le
curseur est deja dans le champ du numero, il tape les derniers chiffres et
appuie sur Entree. Tout le reste — le prefixe, le dossier, le fichier — est
regle une fois pour toutes dans un profil.

Une fois la serie terminee, le bouton du bas bascule sur l'onglet d'ecriture
avec la bonne source deja choisie : les adresses relevees partent directement
dans les fiches Excel, sans fichier intermediaire.
"""

from __future__ import annotations

import os
import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog, ttk

from . import arp, releve, runner, serial, surveillance

COULEUR_ATTENTE = "#5c5c5c"
COULEUR_DETECTE = "#0a7a28"
COULEUR_ALERTE = "#c01c28"
COULEUR_INFO = "#1a5fb4"


class OngletBanc(ttk.Frame):
    def __init__(self, parent, application, source=None):
        super().__init__(parent, padding=10)
        self.app = application
        self.source = source                 # None = vraie table ARP
        self.columnconfigure(1, weight=1)

        self.releves = []
        self.mac_a_enregistrer = None
        self.origine_manuelle = False
        self.boucle = None
        self.surveillance = surveillance.Surveillance()

        self.var_dossier = application.var_dossier          # partagé avec l'onglet Excel
        self.var_csv = application.var_csv                  # partagé avec l'onglet Excel
        self.var_ip = tk.StringVar()
        self.var_prefixe = tk.StringVar()
        self.var_vider = tk.BooleanVar(value=True)
        self.var_numero = tk.StringVar()
        self.var_mac = tk.StringVar(value="—")
        self.var_apercu = tk.StringVar(value="")
        self.var_etat_banc = tk.StringVar(value=surveillance.LIBELLES[surveillance.ARRET])

        self._construire()
        self.var_numero.trace_add("write", lambda *_: self._rafraichir_apercu())
        self.var_csv.trace_add("write", lambda *_: self._charger_liste())

    # ------------------------------------------------------------------ vue
    def _construire(self):
        ligne = 0

        ttk.Label(self, text="Adresse IP surveillée").grid(row=ligne, column=0, sticky="w", pady=3)
        cadre = ttk.Frame(self)
        cadre.grid(row=ligne, column=1, sticky="ew", pady=3, padx=(8, 0))
        cadre.columnconfigure(0, weight=1)
        ttk.Entry(cadre, textvariable=self.var_ip).grid(row=0, column=0, sticky="ew")
        ttk.Button(cadre, text="Tester", width=14,
                   command=self._tester_ip).grid(row=0, column=1, padx=(6, 0))
        ligne += 1

        ttk.Label(self, text="Préfixe du numéro").grid(row=ligne, column=0, sticky="w", pady=3)
        cadre = ttk.Frame(self)
        cadre.grid(row=ligne, column=1, sticky="ew", pady=3, padx=(8, 0))
        cadre.columnconfigure(0, weight=1)
        ttk.Entry(cadre, textvariable=self.var_prefixe).grid(row=0, column=0, sticky="ew")
        ttk.Button(cadre, text="Déduire du dossier", width=20,
                   command=self._deduire_prefixe).grid(row=0, column=1, padx=(6, 0))
        ligne += 1

        ttk.Label(self, text="Fichier liste (CSV)").grid(row=ligne, column=0, sticky="w", pady=3)
        cadre = ttk.Frame(self)
        cadre.grid(row=ligne, column=1, sticky="ew", pady=3, padx=(8, 0))
        cadre.columnconfigure(0, weight=1)
        ttk.Entry(cadre, textvariable=self.var_csv).grid(row=0, column=0, sticky="ew")
        ttk.Button(cadre, text="Parcourir…", width=14,
                   command=self._choisir_csv).grid(row=0, column=1, padx=(6, 0))
        ligne += 1

        ttk.Checkbutton(self, text="Vider le cache ARP entre deux appareils "
                                   "(demande les droits administrateur)",
                        variable=self.var_vider).grid(row=ligne, column=1, sticky="w",
                                                      padx=(8, 0), pady=(2, 6))
        ligne += 1

        # --- banc -----------------------------------------------------------
        banc = ttk.LabelFrame(self, text="Banc de test", padding=10)
        banc.grid(row=ligne, column=0, columnspan=2, sticky="ew", pady=4)
        banc.columnconfigure(1, weight=1)

        commandes = ttk.Frame(banc)
        commandes.grid(row=0, column=0, columnspan=3, sticky="ew", pady=(0, 8))
        self.bouton_marche = ttk.Button(commandes, text="Démarrer la surveillance",
                                        command=self._basculer_surveillance, width=26)
        self.bouton_marche.grid(row=0, column=0)
        ttk.Button(commandes, text="Vider le cache ARP", command=self._vider_cache,
                   width=20).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(commandes, text="Saisir la MAC à la main", command=self._saisie_manuelle,
                   width=24).grid(row=0, column=2, padx=(8, 0))
        self.bouton_admin = ttk.Button(commandes, text="Relancer en administrateur",
                                       command=self._relancer_admin, width=26)
        # affiché seulement si un vidage échoue

        self.etiquette_etat_banc = ttk.Label(banc, textvariable=self.var_etat_banc,
                                             foreground=COULEUR_ATTENTE)
        self.etiquette_etat_banc.grid(row=1, column=0, columnspan=3, sticky="w")

        self.etiquette_mac = ttk.Label(banc, textvariable=self.var_mac,
                                       font=("Consolas", 24, "bold"),
                                       foreground=COULEUR_ATTENTE)
        self.etiquette_mac.grid(row=2, column=0, columnspan=3, sticky="w", pady=(6, 6))

        ttk.Label(banc, text="N° appareil").grid(row=3, column=0, sticky="w")
        self.champ_numero = ttk.Entry(banc, textvariable=self.var_numero,
                                      font=("Consolas", 16), width=10, state="disabled")
        self.champ_numero.grid(row=3, column=1, sticky="w", padx=(8, 0))
        self.champ_numero.bind("<Return>", lambda _: self._enregistrer_appareil())
        self.bouton_valider = ttk.Button(banc, text="Enregistrer (Entrée)",
                                         command=self._enregistrer_appareil,
                                         state="disabled", width=22)
        self.bouton_valider.grid(row=3, column=2, sticky="e")
        ttk.Label(banc, textvariable=self.var_apercu, foreground=COULEUR_INFO).grid(
            row=4, column=1, columnspan=2, sticky="w", padx=(8, 0), pady=(4, 0))
        ligne += 1

        # --- liste ----------------------------------------------------------
        cadre_liste = ttk.LabelFrame(self, text="Liste des relevés", padding=6)
        cadre_liste.grid(row=ligne, column=0, columnspan=2, sticky="nsew", pady=6)
        cadre_liste.columnconfigure(0, weight=1)
        cadre_liste.rowconfigure(0, weight=1)
        self.rowconfigure(ligne, weight=1)
        colonnes = ("numero", "mac", "ip", "horodatage")
        titres = {"numero": "N° de série", "mac": "Adresse MAC",
                  "ip": "IP", "horodatage": "Relevé le"}
        largeurs = {"numero": 190, "mac": 170, "ip": 120, "horodatage": 160}
        self.table = ttk.Treeview(cadre_liste, columns=colonnes, show="headings", height=7)
        for colonne in colonnes:
            self.table.heading(colonne, text=titres[colonne])
            self.table.column(colonne, width=largeurs[colonne], anchor="w")
        self.table.grid(row=0, column=0, sticky="nsew")
        barre = ttk.Scrollbar(cadre_liste, orient="vertical", command=self.table.yview)
        barre.grid(row=0, column=1, sticky="ns")
        self.table.configure(yscrollcommand=barre.set)

        boutons = ttk.Frame(cadre_liste)
        boutons.grid(row=1, column=0, columnspan=2, sticky="ew", pady=(6, 0))
        boutons.columnconfigure(2, weight=1)
        ttk.Button(boutons, text="Supprimer la ligne", command=self._supprimer_ligne,
                   width=20).grid(row=0, column=0)
        ttk.Button(boutons, text="Recharger la liste",
                   command=lambda: self._charger_liste(explicite=True),
                   width=20).grid(row=0, column=1, padx=(8, 0))
        ttk.Button(boutons, text="Écrire ces relevés dans les fiches Excel →",
                   command=self._vers_ecriture, width=42).grid(row=0, column=3, sticky="e")

    # ----------------------------------------------------------- réglages
    def lire_reglages(self):
        return {
            "ip_surveillee": self.var_ip.get().strip(),
            "prefixe_numero": self.var_prefixe.get().strip(),
            "vider_cache_arp": self.var_vider.get(),
        }

    def appliquer_reglages(self, profil):
        if self.surveillance.etat != surveillance.ARRET:
            self._arreter()              # le profil remplace la machine à états
        self.var_ip.set(profil["ip_surveillee"])
        self.var_prefixe.set(profil["prefixe_numero"])
        self.var_vider.set(profil["vider_cache_arp"])
        self.surveillance = surveillance.Surveillance(
            profil["scrutations_stables"], profil["scrutations_absence"])
        self._charger_liste()
        self._peindre_etat()

    def _groupes(self):
        return tuple(self.app.reglages()["groupes_numero"])

    def _intervalle(self):
        return self.app.reglages()["intervalle_ms"]

    # ------------------------------------------------------------------ liste
    def _choisir_csv(self):
        dossier = self.var_dossier.get().strip()
        choix = filedialog.asksaveasfilename(
            title="Fichier de la liste — en désigner un existant ou en créer un",
            defaultextension=".csv",
            initialdir=dossier if os.path.isdir(dossier) else None,
            initialfile=os.path.basename(self.var_csv.get().strip()) or releve.NOM_DEFAUT,
            filetypes=[("Fichier CSV", "*.csv"), ("Tous les fichiers", "*.*")],
            confirmoverwrite=False, parent=self)
        if choix:
            self.var_csv.set(choix)

    def _charger_liste(self, explicite=False):
        """Relit le CSV désigné par le champ « Fichier liste ».

        *explicite* vaut True quand l'opérateur a cliqué sur « Recharger la
        liste » : on lui répond alors toujours, même pour dire qu'il n'y a rien
        à charger. Un bouton qui ne dit rien passe pour cassé. La relecture
        automatique, elle, reste discrète : le champ est relu à chaque frappe,
        et annoncer « introuvable » à chaque caractère tapé n'aiderait personne.
        """
        chemin = self.var_csv.get().strip()

        # Tolérance : un dossier saisi à la place du fichier. On y cherche la
        # liste, plutôt que de renvoyer l'opérateur à ses réglages.
        if chemin and os.path.isdir(chemin):
            candidat = os.path.join(chemin, releve.NOM_DEFAUT)
            if os.path.isfile(candidat):
                self.var_csv.set(candidat)       # la trace relance le chargement
                return
            self.releves = []
            self._rafraichir_table()
            if explicite:
                self.app.dire("%s est un dossier, et il ne contient pas de %s. "
                              "Indiquer le chemin complet du fichier."
                              % (chemin, releve.NOM_DEFAUT))
            return

        if not chemin:
            self.releves = []
            self._rafraichir_table()
            if explicite:
                self.app.dire("Aucun fichier de liste choisi : cliquer sur « Parcourir… » "
                              "en face de « Fichier liste (CSV) » et désigner le fichier.")
            return

        if not os.path.isfile(chemin):
            self.releves = []
            self._rafraichir_table()
            if explicite:
                self.app.dire("Fichier de liste introuvable : %s" % chemin)
            return

        self.releves, message = releve.charger(chemin)
        self._rafraichir_table()
        if message:
            self.app.dire(message)
        elif explicite:
            self.app.dire("%s ne contient aucun relevé." % os.path.basename(chemin))

    def _rafraichir_table(self):
        self.table.delete(*self.table.get_children())
        for index, element in enumerate(self.releves):
            self.table.insert("", "end", iid=str(index),
                              values=(element.numero, element.mac,
                                      element.ip, element.horodatage))
        enfants = self.table.get_children()
        if enfants:
            self.table.see(enfants[-1])

    def _sauver_liste(self):
        chemin = self.var_csv.get().strip()
        if not chemin:
            messagebox.showinfo("Liste", "Choisir d'abord le fichier de la liste (CSV) : "
                                         "sans lui, rien ne peut être enregistré.", parent=self)
            return False
        try:
            releve.enregistrer(chemin, self.releves)
        except OSError as erreur:
            messagebox.showerror("Liste", "Écriture impossible : %s" % erreur, parent=self)
            return False
        return True

    def _supprimer_ligne(self):
        selection = self.table.selection()
        if not selection:
            return
        index = int(selection[0])
        element = self.releves[index]
        if not messagebox.askyesno(
                "Liste", "Retirer %s (%s) de la liste ?" % (element.numero, element.mac),
                parent=self):
            return
        del self.releves[index]
        self._rafraichir_table()
        if self._sauver_liste():
            self.app.dire("%s retiré de la liste." % element.numero)

    def _vers_ecriture(self):
        if not self.releves:
            messagebox.showinfo("Écriture", "La liste est vide : rien à écrire.", parent=self)
            return
        self.app.basculer_vers_ecriture()

    def _deduire_prefixe(self):
        dossier = self.var_dossier.get().strip()
        if not os.path.isdir(dossier):
            messagebox.showinfo("Préfixe",
                                "Choisir d'abord le dossier des fiches, dans l'onglet "
                                "« Écriture dans Excel ».", parent=self)
            return
        fiches = runner.numeros_des_fiches(dossier, self.app.reglages()["sous_dossiers"],
                                           self._groupes())
        prefixe, message = serial.deduire_prefixe(
            [os.path.basename(c) for c in fiches.values()], self._groupes())
        if not prefixe:
            messagebox.showwarning("Préfixe", message, parent=self)
            self.app.dire(message)
            return
        self.var_prefixe.set(prefixe)
        self.app.dire(message)

    # ------------------------------------------------------------ surveillance
    def _source_lire(self, ip):
        return self.source.lire(ip) if self.source is not None else arp.lire(ip)

    def _source_vider(self, ip):
        return self.source.vider(ip) if self.source is not None else arp.vider(ip)

    def _basculer_surveillance(self):
        if self.surveillance.etat == surveillance.ARRET:
            self._demarrer()
        else:
            self._arreter()

    def _demarrer(self):
        ip = self.var_ip.get().strip()
        if not ip:
            messagebox.showinfo("Surveillance", "Saisir l'adresse IP à surveiller.", parent=self)
            return
        if not self.var_csv.get().strip():
            messagebox.showinfo("Surveillance", "Choisir d'abord le fichier de la liste (CSV).",
                                parent=self)
            return
        prefixe, erreur = serial.normaliser_prefixe(self.var_prefixe.get(), self._groupes())
        if erreur:
            messagebox.showwarning("Préfixe", erreur, parent=self)
            return
        self.var_prefixe.set(prefixe)
        self.surveillance.demarrer()
        self.bouton_marche.configure(text="Arrêter la surveillance")
        self._peindre_etat()
        self.app.dire("Surveillance de %s en cours." % ip)
        self._scruter()

    def _arreter(self):
        if self.boucle is not None:
            self.after_cancel(self.boucle)
            self.boucle = None
        self.surveillance.arreter()
        self.bouton_marche.configure(text="Démarrer la surveillance")
        self._armer_saisie(None)
        self._peindre_etat()
        self.app.dire("Surveillance arrêtée.")

    def _scruter(self):
        self.boucle = None
        if self.surveillance.etat == surveillance.ARRET:
            return
        ip = self.var_ip.get().strip()
        evenement = self.surveillance.scruter(self._source_lire(ip))
        if evenement:
            self._traiter(evenement, ip)
        self._peindre_etat()
        self.boucle = self.after(self._intervalle(), self._scruter)

    def _traiter(self, evenement, ip):
        nature, mac = evenement
        if nature == surveillance.DETECTION:
            self._avertir_sonore()
            self._armer_saisie(mac)
            self.app.dire("Appareil détecté : %s" % mac)
        elif nature == surveillance.IDENTIQUE:
            self._armer_saisie(None)
            self.var_mac.set(mac)
            self.etiquette_mac.configure(foreground=COULEUR_ALERTE)
            self.app.dire("MAC identique au relevé précédent : cache ARP non vidé ou "
                          "appareil non remplacé. Débrancher, puis vider le cache ARP.")
        elif nature == surveillance.RETRAIT:
            self._armer_saisie(None)
            if self.var_vider.get():
                succes, message = self._source_vider(ip)
                self.app.dire(message)
                if not succes:
                    self.bouton_admin.grid(row=0, column=3, padx=(8, 0))
            else:
                self.app.dire("Appareil retiré. En attente du suivant…")
        elif nature == surveillance.PERDU:
            self._armer_saisie(None)
            self.app.dire("L'appareil a disparu avant la saisie du numéro.")

    def _armer_saisie(self, mac):
        """Ouvre ou ferme le champ du numéro selon qu'un appareil attend d'être nommé."""
        self.mac_a_enregistrer = mac
        self.origine_manuelle = False
        if mac:
            self.var_mac.set(mac)
            self.etiquette_mac.configure(foreground=COULEUR_DETECTE)
            self.champ_numero.configure(state="normal")
            self.bouton_valider.configure(state="normal")
            self.var_numero.set("")
            self.champ_numero.focus_set()
        else:
            self.var_mac.set("—")
            self.etiquette_mac.configure(foreground=COULEUR_ATTENTE)
            self.champ_numero.configure(state="disabled")
            self.bouton_valider.configure(state="disabled")
            self.var_numero.set("")
            self.var_apercu.set("")

    def _peindre_etat(self):
        etat = self.surveillance.etat
        self.var_etat_banc.set(surveillance.LIBELLES[etat])
        couleurs = {surveillance.ARRET: COULEUR_ATTENTE,
                    surveillance.ATTENTE: COULEUR_INFO,
                    surveillance.DETECTE: COULEUR_DETECTE,
                    surveillance.ATTENTE_RETRAIT: COULEUR_INFO}
        self.etiquette_etat_banc.configure(foreground=couleurs[etat])

    def _avertir_sonore(self):
        try:
            import winsound
            winsound.MessageBeep(winsound.MB_ICONASTERISK)
        except (ImportError, RuntimeError, AttributeError):
            try:
                self.bell()
            except tk.TclError:
                pass

    # --------------------------------------------------------- enregistrement
    def _rafraichir_apercu(self):
        if not self.mac_a_enregistrer:
            self.var_apercu.set("")
            return
        complet, erreur = serial.composer(self.var_prefixe.get(),
                                          self.var_numero.get(), self._groupes())
        self.var_apercu.set(complet if complet else (erreur or ""))

    def _enregistrer_appareil(self):
        mac = self.mac_a_enregistrer
        if not mac:
            return
        complet, erreur = serial.composer(self.var_prefixe.get(),
                                          self.var_numero.get(), self._groupes())
        if erreur:
            messagebox.showwarning("Numéro", erreur, parent=self)
            self.champ_numero.focus_set()
            return

        avertissements = releve.controler(self.releves, complet, mac)
        dossier = self.var_dossier.get().strip()
        if os.path.isdir(dossier):
            fiches = runner.numeros_des_fiches(dossier, self.app.reglages()["sous_dossiers"],
                                               self._groupes())
            if fiches and complet not in fiches:
                avertissements.append(
                    "aucune fiche Excel du dossier ne porte le numéro %s "
                    "(faute de frappe ?)" % complet)
        if avertissements:
            texte = "\n".join("• " + a for a in avertissements)
            if not messagebox.askyesno("Vérification",
                                       "%s\n\nEnregistrer quand même %s ?" % (texte, complet),
                                       parent=self, icon="warning"):
                self.champ_numero.focus_set()
                return

        self.releves.append(releve.Releve(complet, mac, self.var_ip.get().strip()))
        self._rafraichir_table()
        if not self._sauver_liste():
            self.releves.pop()
            self._rafraichir_table()
            return

        manuelle = self.origine_manuelle
        self._armer_saisie(None)
        if not manuelle:
            self.surveillance.enregistrer()
        elif not self.surveillance.ignorer():
            # La machine n'avait rien détecté : on retient quand même cette MAC,
            # pour que le garde-fou « MAC identique » joue au coup suivant.
            self.surveillance.derniere_enregistree = mac
        if self.source is not None and hasattr(self.source, "suivant"):
            self.source.suivant()
        self._peindre_etat()
        self.app.dire("%s → %s enregistré (%d au total). Débrancher l'appareil."
                      % (complet, mac, len(self.releves)))

    def _saisie_manuelle(self):
        """Roue de secours : recopier la MAC affichée par Tftpd32.

        La détection automatique couvre le cas normal ; ce bouton évite d'être
        bloqué le jour où elle ne passe pas (appareil muet en ARP, poste sans
        droits, adresse IP inattendue).
        """
        brut = simpledialog.askstring(
            "Saisie manuelle",
            "Adresse MAC relevée dans Tftpd32 :\n(00:30:D6:4C:6E:05, 00-30-D6-4C-6E-05…)",
            parent=self)
        if not brut:
            return
        mac = arp.normaliser(brut)
        exploitable, raison = arp.mac_exploitable(mac)
        if not exploitable:
            messagebox.showwarning("Saisie manuelle", raison, parent=self)
            return
        self._armer_saisie(mac)
        self.origine_manuelle = True
        self.etiquette_mac.configure(foreground=COULEUR_INFO)
        self.app.dire("MAC saisie à la main : %s — saisir le numéro de l'appareil." % mac)

    # ---------------------------------------------------------------- outils
    def _tester_ip(self):
        ip = self.var_ip.get().strip()
        if not ip:
            return
        mac = self._source_lire(ip)
        if mac:
            self.app.dire("%s répond : %s" % (ip, mac))
        else:
            self.app.dire("Aucune réponse de %s (appareil débranché, IP différente, "
                          "ou pas encore d'entrée ARP)." % ip)

    def _vider_cache(self):
        ip = self.var_ip.get().strip()
        if not ip:
            return
        succes, message = self._source_vider(ip)
        self.app.dire(message)
        if not succes:
            self.bouton_admin.grid(row=0, column=3, padx=(8, 0))

    def _relancer_admin(self):
        if not messagebox.askyesno(
                "Administrateur",
                "Le programme va se relancer avec les droits administrateur.\n"
                "Fermer cette fenêtre ensuite.\n\nContinuer ?", parent=self):
            return
        succes, message = arp.relancer_en_administrateur()
        self.app.dire(message)

    def fermer(self):
        if self.boucle is not None:
            self.after_cancel(self.boucle)
            self.boucle = None
        return True
