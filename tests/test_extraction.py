"""Tests : python -m unittest discover -s tests"""

from __future__ import annotations

import json
import os
import shutil
import sys
import tempfile
import unittest
import xml.dom.minidom
import zipfile

from extraction_mac import config, docmac, runner, serial, xlsxcell
from tests import fixtures


class TestNumeroDeSerie(unittest.TestCase):
    def test_formes_du_nom_de_fichier(self):
        attendu = "260918-0144215-00092"
        for nom in (
            "X130392_B_260918-0144215-00092.doc",
            "X130392_B_260918014421500092.doc",
            "X301523-9_MF19-Ecran-Cabine-12.1_Fiche-de-Test N°260918-0144215-00092.xlsx",
            "260918_0144215_00092.docx",
            "copie de 260918-0144215-00092 (1).xlsx",
        ):
            with self.subTest(nom=nom):
                self.assertEqual(serial.key_from_filename(nom), (attendu, None))

    def test_numero_absent(self):
        numero, erreur = serial.key_from_filename("fiche_sans_numero.doc")
        self.assertIsNone(numero)
        self.assertIn("aucun numéro", erreur)

    def test_deux_numeros_differents(self):
        numero, erreur = serial.key_from_filename("260918-0144215-00092_et_260918-0144215-00093.doc")
        self.assertIsNone(numero)
        self.assertIn("plusieurs", erreur)

    def test_pas_de_faux_positif_au_milieu_des_chiffres(self):
        self.assertEqual(serial.find_keys("9999260918014421500092999"), [])

    def test_longueurs_de_groupes_personnalisees(self):
        self.assertEqual(serial.key_from_filename("AB_1234-567.docx", groups=(4, 3)),
                         ("1234-567", None))


class TestAdresseMac(unittest.TestCase):
    def setUp(self):
        self.dossier = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.dossier, True)

    def _docx(self, lignes, nom="essai.docx"):
        return fixtures.ecrire_docx(os.path.join(self.dossier, nom), lignes)

    def test_mac_avec_tirets_devient_deux_points(self):
        chemin = self._docx(["Numéro de série : X130392_B_260918-0144215-00092",
                             "Adresse MAC : 00-30-D6-4C-6E-05",
                             "Nom de l'opérateur : LG"])
        self.assertEqual(docmac.extraire_mac(chemin), ("00:30:D6:4C:6E:05", None))

    def test_mac_deja_en_deux_points(self):
        chemin = self._docx(["Adresse MAC : 00:30:d6:4c:6e:05"])
        self.assertEqual(docmac.extraire_mac(chemin), ("00:30:D6:4C:6E:05", None))

    def test_mac_collee_acceptee_si_libelle_present(self):
        chemin = self._docx(["Adresse MAC : 0030D64C6E05"])
        self.assertEqual(docmac.extraire_mac(chemin), ("00:30:D6:4C:6E:05", None))

    def test_numero_de_serie_jamais_pris_pour_une_mac(self):
        chemin = self._docx(["Numéro de série : X130392_B_260918-0144215-00092",
                             "Référence : 260902-0144069-00092"])
        mac, erreur = docmac.extraire_mac(chemin)
        self.assertIsNone(mac)
        self.assertIn("aucune adresse MAC", erreur)

    def test_deux_mac_differentes_refusees(self):
        chemin = self._docx(["MAC eth0 : 00:11:22:33:44:55", "MAC eth1 : 00:11:22:33:44:66"])
        mac, erreur = docmac.extraire_mac(chemin)
        self.assertIsNone(mac)
        self.assertIn("plusieurs adresses MAC", erreur)

    def test_mac_nulle_refusee(self):
        chemin = self._docx(["Adresse MAC : 00-00-00-00-00-00"])
        mac, erreur = docmac.extraire_mac(chemin)
        self.assertIsNone(mac)
        self.assertIn("invalide", erreur)

    def test_doc_binaire_word_97(self):
        chemin = os.path.join(self.dossier, "ancien.doc")
        with open(chemin, "wb") as fichier:
            fichier.write(b"\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1" + b"\x00" * 100)
            fichier.write("Adresse MAC : 00-30-D6-4C-6E-05".encode("utf-16-le"))
            fichier.write(b"\x00" * 50)
        self.assertEqual(docmac.extraire_mac(chemin), ("00:30:D6:4C:6E:05", None))

    def test_separateur_et_casse_configurables(self):
        chemin = self._docx(["Adresse MAC : 00-30-D6-4C-6E-05"])
        self.assertEqual(docmac.extraire_mac(chemin, "-", False), ("00-30-d6-4c-6e-05", None))


class TestEcritureExcel(unittest.TestCase):
    def setUp(self):
        self.dossier = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.dossier, True)
        self.classeur = fixtures.ecrire_xlsx(os.path.join(self.dossier, "fiche.xlsx"))

    def _pieces(self, chemin):
        with zipfile.ZipFile(chemin) as zf:
            return {info.filename: zf.read(info.filename) for info in zf.infolist()}

    def test_ecriture_et_relecture(self):
        self.assertIsNone(xlsxcell.lire_cellule(self.classeur, "Constit produit", "F27"))
        xlsxcell.ecrire_cellule(self.classeur, "Constit produit", "F27", "00:30:D6:4C:6E:05")
        self.assertEqual(xlsxcell.lire_cellule(self.classeur, "Constit produit", "F27"),
                         "00:30:D6:4C:6E:05")

    def test_seule_la_feuille_visee_est_modifiee(self):
        avant = self._pieces(self.classeur)
        xlsxcell.ecrire_cellule(self.classeur, "Constit produit", "F27", "00:30:D6:4C:6E:05")
        apres = self._pieces(self.classeur)
        self.assertEqual(list(avant), list(apres))
        modifiees = [nom for nom in avant if avant[nom] != apres[nom]]
        self.assertEqual(modifiees, ["xl/worksheets/sheet1.xml"])

    def test_style_de_la_cellule_conserve(self):
        xlsxcell.ecrire_cellule(self.classeur, "Constit produit", "F27", "00:30:D6:4C:6E:05")
        with zipfile.ZipFile(self.classeur) as zf:
            feuille = zf.read("xl/worksheets/sheet1.xml").decode()
        self.assertIn('<c r="F27" s="98" t="inlineStr">', feuille)

    def test_xml_reste_valide(self):
        xlsxcell.ecrire_cellule(self.classeur, "Constit produit", "F27", "a & b < c")
        with zipfile.ZipFile(self.classeur) as zf:
            for info in zf.infolist():
                if info.filename.endswith((".xml", ".rels")):
                    xml.dom.minidom.parseString(zf.read(info.filename))
        self.assertEqual(xlsxcell.lire_cellule(self.classeur, "Constit produit", "F27"), "a & b < c")

    def test_cellule_absente_est_creee_au_bon_rang(self):
        xlsxcell.ecrire_cellule(self.classeur, "Constit produit", "D27", "valeur")
        with zipfile.ZipFile(self.classeur) as zf:
            feuille = zf.read("xl/worksheets/sheet1.xml").decode()
        self.assertLess(feuille.index('r="B27"'), feuille.index('r="D27"'))
        self.assertLess(feuille.index('r="D27"'), feuille.index('r="F27"'))
        self.assertEqual(xlsxcell.lire_cellule(self.classeur, "Constit produit", "D27"), "valeur")

    def test_ligne_absente_est_creee_au_bon_rang(self):
        classeur = fixtures.ecrire_xlsx(
            os.path.join(self.dossier, "trous.xlsx"),
            lignes={10: [("A10", 0, None)], 30: [("A30", 1, None)]})
        xlsxcell.ecrire_cellule(classeur, "Constit produit", "C20", "milieu")
        with zipfile.ZipFile(classeur) as zf:
            feuille = zf.read("xl/worksheets/sheet1.xml").decode()
        self.assertLess(feuille.index('r="10"'), feuille.index('r="20"'))
        self.assertLess(feuille.index('r="20"'), feuille.index('r="30"'))
        self.assertEqual(xlsxcell.lire_cellule(classeur, "Constit produit", "C20"), "milieu")

    def test_cellule_fusionnee_ecrit_dans_l_ancre(self):
        classeur = fixtures.ecrire_xlsx(
            os.path.join(self.dossier, "fusion.xlsx"),
            lignes={5: [("B5", None, None)]}, fusions=("B5:E5",))
        xlsxcell.ecrire_cellule(classeur, "Constit produit", "D5", "00:30:D6:4C:6E:05")
        self.assertEqual(xlsxcell.lire_cellule(classeur, "Constit produit", "B5"),
                         "00:30:D6:4C:6E:05")

    def test_feuille_absente(self):
        with self.assertRaises(xlsxcell.ErreurExcel) as contexte:
            xlsxcell.ecrire_cellule(self.classeur, "Onglet inconnu", "F27", "x")
        self.assertIn("absente du classeur", str(contexte.exception))

    def test_reference_invalide(self):
        with self.assertRaises(xlsxcell.ErreurExcel):
            xlsxcell.normaliser_ref("27F")

    def test_reference_tolerante(self):
        self.assertEqual(xlsxcell.normaliser_ref(" $f$27 "), "F27")

    def test_lister_feuilles_et_libelles(self):
        self.assertEqual(xlsxcell.lister_feuilles(self.classeur), ["Constit produit"])
        self.assertEqual(xlsxcell.trouver_libelles(self.classeur, "MAC"),
                         [("Constit produit", "B27", "@MAC")])

    def test_xls_refuse(self):
        chemin = os.path.join(self.dossier, "ancien.xls")
        open(chemin, "wb").close()
        with self.assertRaises(xlsxcell.ErreurExcel) as contexte:
            xlsxcell.lister_feuilles(chemin)
        self.assertIn("non géré", str(contexte.exception))


class TestChaineComplete(unittest.TestCase):
    def setUp(self):
        self.dossier = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.dossier, True)

    def _paire(self, numero, mac="00-30-D6-4C-6E-05"):
        fixtures.ecrire_docx(os.path.join(self.dossier, "X130392_B_%s.doc" % numero),
                             ["Numéro de série : X130392_B_%s" % numero, "Adresse MAC : %s" % mac])
        fixtures.ecrire_xlsx(os.path.join(self.dossier, "X301523-9_Fiche-de-Test N°%s.xlsx" % numero))

    def _options(self, **extra):
        base = dict(dossier=self.dossier, feuille="Constit produit", cellule="F27")
        base.update(extra)
        return runner.Options(**base)

    def test_simulation_puis_ecriture_puis_relance(self):
        self._paire("260918-0144215-00092")
        bilan = runner.executer(self._options(simulation=True))
        self.assertEqual(bilan.compter(runner.SIMULATION), 1)
        classeur = os.path.join(self.dossier, "X301523-9_Fiche-de-Test N°260918-0144215-00092.xlsx")
        self.assertIsNone(xlsxcell.lire_cellule(classeur, "Constit produit", "F27"))

        bilan = runner.executer(self._options())
        self.assertEqual(bilan.compter(runner.ECRIT), 1)
        self.assertEqual(xlsxcell.lire_cellule(classeur, "Constit produit", "F27"),
                         "00:30:D6:4C:6E:05")

        bilan = runner.executer(self._options())
        self.assertEqual(bilan.compter(runner.DEJA_OK), 1)

    def test_valeur_differente_protegee_sans_ecrasement(self):
        self._paire("260918-0144215-00092")
        classeur = os.path.join(self.dossier, "X301523-9_Fiche-de-Test N°260918-0144215-00092.xlsx")
        xlsxcell.ecrire_cellule(classeur, "Constit produit", "F27", "AA:BB:CC:DD:EE:FF")

        bilan = runner.executer(self._options())
        self.assertEqual(bilan.compter(runner.IGNORE), 1)
        self.assertEqual(xlsxcell.lire_cellule(classeur, "Constit produit", "F27"),
                         "AA:BB:CC:DD:EE:FF")

        bilan = runner.executer(self._options(ecraser=True))
        self.assertEqual(bilan.compter(runner.ECRIT), 1)
        self.assertEqual(xlsxcell.lire_cellule(classeur, "Constit produit", "F27"),
                         "00:30:D6:4C:6E:05")

    def test_appairage_de_plusieurs_paires(self):
        for numero in ("260918-0144215-00092", "260918-0144215-00093", "260918-0144215-00094"):
            self._paire(numero)
        bilan = runner.executer(self._options())
        self.assertEqual(bilan.compter(runner.ECRIT), 3)
        self.assertEqual(bilan.compter(runner.ERREUR), 0)

    def test_word_orphelin_et_excel_orphelin(self):
        self._paire("260918-0144215-00092")
        fixtures.ecrire_docx(os.path.join(self.dossier, "X130392_B_260918-0144215-00093.doc"),
                             ["Adresse MAC : 00-11-22-33-44-55"])
        fixtures.ecrire_xlsx(os.path.join(self.dossier, "Fiche N°260918-0144215-00094.xlsx"))
        bilan = runner.executer(self._options())
        self.assertEqual(bilan.compter(runner.ECRIT), 1)
        self.assertEqual(bilan.compter(runner.IGNORE), 2)

    def test_fichier_temporaire_excel_ignore(self):
        self._paire("260918-0144215-00092")
        with open(os.path.join(self.dossier, "~$X301523-9_Fiche-de-Test N°260918-0144215-00092.xlsx"),
                  "wb") as fichier:
            fichier.write(b"verrou")
        bilan = runner.executer(self._options())
        self.assertEqual(bilan.compter(runner.ECRIT), 1)
        self.assertEqual(bilan.compter(runner.ERREUR), 0)

    def test_doublon_excel_signale_sans_ecrire(self):
        self._paire("260918-0144215-00092")
        fixtures.ecrire_xlsx(os.path.join(self.dossier, "copie N°260918-0144215-00092.xlsx"))
        bilan = runner.executer(self._options())
        self.assertEqual(bilan.compter(runner.ERREUR), 1)
        self.assertEqual(bilan.compter(runner.ECRIT), 0)

    def test_sous_dossiers(self):
        sous = os.path.join(self.dossier, "lot 42")
        os.makedirs(sous)
        fixtures.ecrire_docx(os.path.join(sous, "X130392_B_260918-0144215-00095.doc"),
                             ["Adresse MAC : 00-11-22-33-44-55"])
        fixtures.ecrire_xlsx(os.path.join(sous, "Fiche N°260918-0144215-00095.xlsx"))
        self.assertEqual(runner.executer(self._options()).compter(runner.ECRIT), 0)
        self.assertEqual(runner.executer(self._options(sous_dossiers=True)).compter(runner.ECRIT), 1)

    def test_dossier_inexistant(self):
        with self.assertRaises(ValueError):
            runner.executer(self._options(dossier=os.path.join(self.dossier, "absent")))


class TestConfiguration(unittest.TestCase):
    def setUp(self):
        self.dossier = tempfile.mkdtemp()
        self.addCleanup(shutil.rmtree, self.dossier, True)
        self.repli = os.path.join(self.dossier, "repli")
        self.programme = os.path.join(self.dossier, "programme")
        os.makedirs(self.programme)
        self._ancien = dict(os.environ)
        os.environ["APPDATA"] = self.repli
        os.environ["XDG_CONFIG_HOME"] = self.repli
        self.addCleanup(lambda: (os.environ.clear(), os.environ.update(self._ancien)))
        self._vrai_dossier_programme = config.dossier_programme
        config.dossier_programme = lambda: self.programme
        self.addCleanup(setattr, config, "dossier_programme", self._vrai_dossier_programme)

    def test_aller_retour(self):
        profil = config.profil_vide()
        profil.update(dossier="D:\\Fiches", feuille="Constit produit", cellule="F27", ecraser=True)
        config.enregistrer({"MF19": profil}, "MF19")
        profils, dernier = config.charger()
        self.assertEqual(dernier, "MF19")
        self.assertEqual(profils["MF19"]["cellule"], "F27")
        self.assertTrue(profils["MF19"]["ecraser"])

    def test_fichier_illisible_ne_plante_pas(self):
        os.makedirs(config.dossier_config(), exist_ok=True)
        with open(config.chemin_config(), "w", encoding="utf-8") as fichier:
            fichier.write("{ceci n'est pas du json")
        self.assertEqual(config.charger(), ({}, ""))

    def test_profil_incomplet_complete_par_defaut(self):
        config.enregistrer({"partiel": {"cellule": "B12"}}, "partiel")
        profils, _ = config.charger()
        self.assertEqual(profils["partiel"]["cellule"], "B12")
        self.assertEqual(profils["partiel"]["separateur_mac"], ":")
        self.assertEqual(profils["partiel"]["groupes_numero"], [6, 7, 5])

    def test_ecrit_a_cote_du_programme(self):
        destination = config.enregistrer({"MF19": config.profil_vide()}, "MF19")
        attendu = os.path.join(self.programme, config.NOM_FICHIER)
        self.assertEqual(destination, attendu)
        self.assertTrue(os.path.isfile(attendu))
        self.assertFalse(os.path.exists(os.path.join(self.repli, config.NOM_FICHIER)))

    def test_repli_quand_le_dossier_du_programme_est_inaccessible(self):
        # un chemin situe « sous » un fichier ordinaire ne peut jamais etre cree,
        # y compris pour l'administrateur : cela simule un dossier non inscriptible.
        bloqueur = os.path.join(self.dossier, "fichier.txt")
        with open(bloqueur, "w", encoding="utf-8") as fichier:
            fichier.write("x")
        config.dossier_programme = lambda: os.path.join(bloqueur, "sous-dossier")

        destination = config.enregistrer({"MF19": config.profil_vide()}, "MF19")
        self.assertEqual(destination, os.path.join(config.dossier_repli(), config.NOM_FICHIER))
        self.assertTrue(os.path.isfile(destination))
        profils, dernier = config.charger()
        self.assertEqual(dernier, "MF19")

    def test_le_fichier_a_cote_du_programme_est_prioritaire(self):
        os.makedirs(config.dossier_repli(), exist_ok=True)
        for dossier, cellule in ((config.dossier_repli(), "Z99"), (self.programme, "F27")):
            with open(os.path.join(dossier, config.NOM_FICHIER), "w", encoding="utf-8") as fichier:
                json.dump({"profils": {"MF19": {"cellule": cellule}},
                           "dernier_profil": "MF19"}, fichier)
        profils, _ = config.charger()
        self.assertEqual(profils["MF19"]["cellule"], "F27")

    def test_dossier_programme_suit_l_executable_si_fige(self):
        config.dossier_programme = self._vrai_dossier_programme
        faux_exe = os.path.join(self.dossier, "ailleurs", "ExtractionMAC.exe")
        os.makedirs(os.path.dirname(faux_exe))
        anciens = getattr(sys, "frozen", None), sys.executable
        sys.frozen, sys.executable = True, faux_exe
        try:
            self.assertEqual(config.dossier_programme(), os.path.dirname(faux_exe))
        finally:
            sys.executable = anciens[1]
            if anciens[0] is None:
                del sys.frozen
            else:
                sys.frozen = anciens[0]


if __name__ == "__main__":
    unittest.main()
