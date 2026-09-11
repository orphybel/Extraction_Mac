"""Tests de la fenêtre elle-même.

Ces tests existent à cause d'une panne réelle : une méthode nommée ``_options``
dans un onglet masquait ``tkinter.Misc._options``, que ``BaseWidget.__init__``
appelle pour convertir les options d'un widget. Le programme plantait donc au
lancement — et rien ne l'avait vu, parce qu'aucun test ne construisait la
fenêtre et que le contrôle d'intégration se contentait de vérifier que le
processus vivait encore (or, en mode fenêtre, PyInstaller affiche une boîte
d'erreur et le processus reste vivant).

Deux niveaux de protection :

* ``TestCollisionsTkinter`` n'a besoin d'aucun affichage et attrape toute la
  famille du bug, pas seulement ce nom-là ;
* ``TestFenetre`` construit réellement l'interface et la fait travailler ; c'est
  le seul test qui traverse l'interface de bout en bout.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import time
import unittest

try:
    import tkinter as tk
    from tkinter import ttk
    TKINTER = True
except ImportError:                                # poste sans Tkinter
    TKINTER = False


def _affichage_disponible():
    """Tk peut-il réellement ouvrir une fenêtre ici ?"""
    if not TKINTER:
        return False
    try:
        racine = tk.Tk()
    except Exception:                              # noqa: BLE001 - TclError et consorts
        return False
    racine.destroy()
    return True


AFFICHAGE = _affichage_disponible()

if TKINTER:
    from extraction_mac import config, gui, gui_banc, gui_excel, releve
    from . import fixtures


@unittest.skipUnless(TKINTER, "Tkinter n'est pas installé")
class TestCollisionsTkinter(unittest.TestCase):
    """Aucune méthode d'onglet ne doit porter le nom d'un membre de ttk.Frame.

    C'est le garde-fou générique : il ne demande pas d'affichage, donc il tourne
    partout, et il vaut pour les 214 membres hérités — pas seulement pour celui
    qui nous a coûté une version.
    """

    def _classes(self):
        return [gui.Application, gui_excel.OngletExcel, gui_banc.OngletBanc]

    def test_aucune_methode_ne_masque_un_membre_herite(self):
        herites = set(dir(ttk.Frame))
        for classe in self._classes():
            propres = {nom for nom, valeur in vars(classe).items()
                       if callable(valeur) or isinstance(valeur, property)}
            collisions = sorted(propres & herites - {"__init__", "__module__",
                                                     "__qualname__", "__doc__"})
            self.assertEqual(
                collisions, [],
                "%s redéfinit %s, qui appartient déjà à un widget Tkinter : "
                "tkinter l'appellerait à notre place." % (classe.__name__, collisions))

    def test_le_nom_fautif_est_bien_detecte(self):
        """Contrôle négatif : le garde-fou doit réagir au bug qu'il vise."""
        herites = set(dir(ttk.Frame))
        self.assertIn("_options", herites)
        self.assertNotIn("_options", vars(gui_excel.OngletExcel))
        self.assertIn("_options_runner", vars(gui_excel.OngletExcel))


class _FauxMessagebox:
    """Remplace les boîtes de dialogue : une modale bloquerait le test pour toujours."""

    def __init__(self):
        self.appels = []

    def showinfo(self, titre, message, **_):
        self.appels.append(("info", titre, message))

    def showwarning(self, titre, message, **_):
        self.appels.append(("avertissement", titre, message))

    def showerror(self, titre, message, **_):
        self.appels.append(("erreur", titre, message))

    def askyesno(self, titre, message, **_):
        self.appels.append(("question", titre, message))
        return True

    def erreurs(self):
        return [a for a in self.appels if a[0] == "erreur"]


@unittest.skipUnless(AFFICHAGE, "aucun affichage disponible pour Tk")
class TestFenetre(unittest.TestCase):
    def setUp(self):
        self.dossier = tempfile.mkdtemp()

        # la configuration doit rester dans le bac à sable
        self._vrai_dossier_config = config.dossier_config
        self._vrais_chemins = config.chemins_lecture
        config.dossier_config = lambda: self.dossier
        config.chemins_lecture = lambda: [os.path.join(self.dossier, config.NOM_FICHIER)]

        self.boites = _FauxMessagebox()
        self._vraies_boites = {}
        for module in (gui, gui_excel, gui_banc):
            self._vraies_boites[module] = module.messagebox
            module.messagebox = self.boites

        self.racine = tk.Tk()
        self.racine.withdraw()                     # ne pas faire clignoter de fenêtre
        self.app = gui.Application(self.racine, simulation=True)
        self.racine.update()

    def tearDown(self):
        try:
            self.app._fermer()          # le vrai chemin de fermeture
        except tk.TclError:
            pass
        for module, vraie in self._vraies_boites.items():
            module.messagebox = vraie
        config.dossier_config = self._vrai_dossier_config
        config.chemins_lecture = self._vrais_chemins
        shutil.rmtree(self.dossier, ignore_errors=True)

    # ------------------------------------------------------------ structure
    def test_la_fenetre_se_construit(self):
        """Le test qui manquait : il échouait avec le bug ``_options``."""
        self.assertIsNotNone(self.app.onglet_banc)
        self.assertIsNotNone(self.app.onglet_excel)
        self.assertEqual(self.boites.erreurs(), [])

    def test_les_deux_onglets_sont_presents(self):
        titres = [self.app.onglets.tab(i, "text").strip()
                  for i in range(len(self.app.onglets.tabs()))]
        self.assertEqual(titres, ["Relevé au banc", "Écriture dans Excel"])

    def test_reglages_fusionne_les_deux_onglets(self):
        self.app.onglet_excel.var_cellule.set("H12")
        self.app.onglet_banc.var_ip.set("10.0.0.5")
        self.app.var_dossier.set(self.dossier)
        reglages = self.app.reglages()
        self.assertEqual(reglages["cellule"], "H12")
        self.assertEqual(reglages["ip_surveillee"], "10.0.0.5")
        self.assertEqual(reglages["dossier"], self.dossier)

    def test_le_dossier_est_partage_par_les_deux_onglets(self):
        self.app.onglet_excel.var_dossier.set(self.dossier)
        self.assertEqual(self.app.onglet_banc.var_dossier.get(), self.dossier)

    def test_aller_retour_d_un_profil(self):
        self.app.var_dossier.set(self.dossier)
        self.app.onglet_excel.var_cellule.set("H12")
        self.app.onglet_banc.var_prefixe.set("260918-0144215")
        self.app.var_profil.set("banc 1")
        self.app._enregistrer_profil()

        self.app.onglet_excel.var_cellule.set("A1")
        self.app.onglet_banc.var_prefixe.set("")
        self.app._charger_profil()

        self.assertEqual(self.app.onglet_excel.var_cellule.get(), "H12")
        self.assertEqual(self.app.onglet_banc.var_prefixe.get(), "260918-0144215")

    def test_bascule_vers_l_ecriture(self):
        self.app.basculer_vers_ecriture()
        self.racine.update()
        self.assertEqual(self.app.onglet_excel.var_source.get(), gui_excel.SOURCE_RELEVE)
        self.assertEqual(self.racine.nametowidget(self.app.onglets.select()),
                         self.app.onglet_excel)

    def test_detection_simulee_arme_la_saisie(self):
        """Détection → champ ouvert → aperçu du numéro complet, depuis la fenêtre."""
        banc = self.app.onglet_banc
        banc.var_ip.set("192.168.0.100")
        banc.var_prefixe.set("260918-0144215")
        self.app.var_csv.set(os.path.join(self.dossier, "liste.csv"))
        banc._demarrer()
        for _ in range(4):
            banc._scruter()
        self.racine.update()

        self.assertEqual(str(banc.champ_numero.cget("state")), "normal")
        self.assertEqual(banc.var_mac.get(), banc.source.macs[0])

        banc.var_numero.set("92")
        self.racine.update()
        self.assertEqual(banc.var_apercu.get(), "260918-0144215-00092")
        banc._arreter()

    def test_enregistrement_depuis_la_fenetre(self):
        """Saisie validée → ligne dans le tableau et dans le CSV."""
        banc = self.app.onglet_banc
        csv = os.path.join(self.dossier, "liste.csv")
        banc.var_ip.set("192.168.0.100")
        banc.var_prefixe.set("260918-0144215")
        self.app.var_csv.set(csv)
        banc._demarrer()
        for _ in range(4):
            banc._scruter()
        self.racine.update()

        attendue = banc.var_mac.get()
        banc.var_numero.set("92")
        banc._enregistrer_appareil()
        self.racine.update()
        banc._arreter()

        self.assertEqual(len(banc.table.get_children()), 1)
        relus, _ = releve.charger(csv)
        self.assertEqual([(r.numero, r.mac) for r in relus],
                         [("260918-0144215-00092", attendue)])

    # ------------------------------------------------------ traitement réel
    def _preparer_dossier(self):
        numero, mac = "260918-0144215-00092", "00:30:D6:4C:6E:05"
        fixtures.ecrire_xlsx(os.path.join(self.dossier, "Fiche N%s.xlsx" % numero))
        self.app.var_dossier.set(self.dossier)
        self.app.onglet_excel.var_feuille.set("Constit produit")
        self.app.onglet_excel.var_cellule.set("F27")
        return numero, mac

    def _attendre_bilan(self, secondes=20):
        journal = self.app.onglet_excel.journal
        limite = time.time() + secondes
        while time.time() < limite:
            self.racine.update()
            if "Bilan :" in journal.get("1.0", "end"):
                return journal.get("1.0", "end")
            time.sleep(0.05)
        self.fail("le traitement ne s'est pas terminé :\n%s" % journal.get("1.0", "end"))

    def test_analyse_depuis_les_fichiers_word(self):
        numero, mac = self._preparer_dossier()
        fixtures.ecrire_docx(os.path.join(self.dossier, "X130392_%s.doc" % numero),
                             ["Adresse MAC : %s" % mac])
        self.app.onglet_excel.var_source.set(gui_excel.SOURCE_WORD)
        self.app.onglet_excel._lancer(True)
        texte = self._attendre_bilan()
        self.assertIn("À ÉCRIRE", texte)
        self.assertIn(mac, texte)
        self.assertEqual(self.boites.erreurs(), [])

    def test_analyse_depuis_la_liste_relevee(self):
        numero, mac = self._preparer_dossier()
        csv = os.path.join(self.dossier, "MAC-releves.csv")
        releve.enregistrer(csv, [releve.Releve(numero, mac, "192.168.0.100")])
        self.app.var_csv.set(csv)
        self.app.onglet_excel.var_source.set(gui_excel.SOURCE_RELEVE)
        self.app.onglet_excel._lancer(True)
        texte = self._attendre_bilan()
        self.assertIn("À ÉCRIRE", texte)
        self.assertIn("liste relevée au banc", texte)
        self.assertEqual(self.boites.erreurs(), [])

    def test_liste_relevee_absente_signalee_sans_planter(self):
        self._preparer_dossier()
        self.app.var_csv.set("")
        self.app.onglet_excel.var_source.set(gui_excel.SOURCE_RELEVE)
        self.app.onglet_excel._lancer(True)
        self.racine.update()
        self.assertTrue(self.boites.erreurs(), "l'absence de liste doit être signalée")

    # ------------------------------------------- rechargement d'une liste
    def _ecrire_liste_v1(self, dossier=None, nom=None):
        """Un CSV tel que la version précédente l'écrivait."""
        chemin = os.path.join(dossier or self.dossier, nom or releve.NOM_DEFAUT)
        with open(chemin, "w", encoding="utf-8-sig", newline="") as fichier:
            fichier.write("numero;mac;ip;horodatage\r\n"
                          "260918-0144215-00092;00:30:D6:4C:6E:05;192.168.0.100;"
                          "2026-09-10 08:12:03\r\n"
                          "260918-0144215-00093;0A:1B:2C:3D:4E:5F;192.168.0.100;"
                          "2026-09-10 08:19:44\r\n")
        return chemin

    def test_liste_v1_rechargee_pour_etre_continuee(self):
        """Le cas réel : reprendre une liste commencée avec la version précédente."""
        chemin = self._ecrire_liste_v1()
        self.app.var_csv.set(chemin)
        self.app.onglet_banc._charger_liste(explicite=True)
        self.racine.update()
        self.assertEqual(len(self.app.onglet_banc.table.get_children()), 2)
        self.assertEqual([r.numero for r in self.app.onglet_banc.releves],
                         ["260918-0144215-00092", "260918-0144215-00093"])
        self.assertIn("2 relevé", self.app.var_etat.get())

    def test_designer_le_dossier_suffit_a_trouver_la_liste(self):
        """Ce que l'on attend spontanément : renseigner le dossier des PV."""
        self._ecrire_liste_v1()
        self.app.var_csv.set("")
        self.app.var_dossier.set(self.dossier)
        self.racine.update()
        self.assertEqual(self.app.var_csv.get(),
                         os.path.join(self.dossier, releve.NOM_DEFAUT))
        self.assertEqual(len(self.app.onglet_banc.releves), 2)

    def test_un_chemin_choisi_a_la_main_n_est_jamais_ecrase(self):
        autre = self._ecrire_liste_v1(nom="ma-liste.csv")
        self.app.var_csv.set(autre)
        self.app.var_dossier.set(self.dossier)
        self.racine.update()
        self.assertEqual(self.app.var_csv.get(), autre)

    def test_recharger_sans_fichier_le_dit(self):
        """Le défaut signalé : le bouton restait muet."""
        self.app.var_csv.set("")
        self.app.onglet_banc._charger_liste(explicite=True)
        self.racine.update()
        self.assertIn("Aucun fichier de liste choisi", self.app.var_etat.get())

    def test_recharger_un_fichier_absent_le_dit(self):
        manquant = os.path.join(self.dossier, "pas-la.csv")
        self.app.var_csv.set(manquant)
        self.app.onglet_banc._charger_liste(explicite=True)
        self.racine.update()
        self.assertIn("introuvable", self.app.var_etat.get())

    def test_un_dossier_saisi_a_la_place_du_fichier_est_toleré(self):
        self._ecrire_liste_v1()
        self.app.var_csv.set(self.dossier)
        self.app.onglet_banc._charger_liste(explicite=True)
        self.racine.update()
        self.assertEqual(self.app.var_csv.get(),
                         os.path.join(self.dossier, releve.NOM_DEFAUT))
        self.assertEqual(len(self.app.onglet_banc.releves), 2)

    def test_la_relecture_automatique_reste_discrete(self):
        """Le champ est relu à chaque frappe : pas de message à chaque caractère."""
        self.app.dire("état de départ")
        self.app.var_csv.set(os.path.join(self.dossier, "pas-la.csv"))
        self.racine.update()
        self.assertEqual(self.app.var_etat.get(), "état de départ")

    def test_une_liste_rechargee_peut_etre_continuee(self):
        """Reprise : un appareil ajouté s'écrit à la suite, sans perdre les anciens."""
        chemin = self._ecrire_liste_v1()
        banc = self.app.onglet_banc
        self.app.var_csv.set(chemin)
        banc._charger_liste(explicite=True)
        banc.var_ip.set("192.168.0.100")
        banc.var_prefixe.set("260918-0144215")
        banc._demarrer()
        for _ in range(4):
            banc._scruter()
        self.racine.update()
        banc.var_numero.set("94")
        banc._enregistrer_appareil()
        self.racine.update()
        banc._arreter()

        relus, _ = releve.charger(chemin)
        self.assertEqual([r.numero for r in relus],
                         ["260918-0144215-00092", "260918-0144215-00093",
                          "260918-0144215-00094"])

    def test_fermeture_propre(self):
        self.app.onglet_banc.var_ip.set("192.168.0.100")
        self.app.var_csv.set(os.path.join(self.dossier, "liste.csv"))
        self.app.onglet_banc.var_prefixe.set("260918-0144215")
        self.app.onglet_banc._demarrer()
        self.racine.update()
        self.assertIsNotNone(self.app.onglet_banc.boucle)
        self.assertTrue(self.app.onglet_banc.fermer())
        self.assertIsNone(self.app.onglet_banc.boucle)


if __name__ == "__main__":
    unittest.main()
