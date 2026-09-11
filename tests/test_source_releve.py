"""La MAC relevée au banc va-t-elle dans la bonne cellule de la bonne fiche ?

Ces tests remplacent — et améliorent — la vérification qui existait entre les
deux anciens dépôts : elle exigeait de cloner l'autre programme et n'était donc
lancée qu'à la main. Maintenant que tout tient dans un seul projet, elle est un
test ordinaire, exécuté à chaque fois.

Chaque cas va jusqu'au bout : on relit la cellule du classeur écrit.
"""

from __future__ import annotations

import os
import shutil
import tempfile
import unittest

from extraction_mac import releve, runner, xlsxcell

from . import fixtures

FEUILLE, CELLULE = "Constit produit", "F27"

N1, M1 = "260918-0144215-00092", "00:30:D6:4C:6E:05"
N2, M2 = "260918-0144215-00093", "0A:1B:2C:3D:4E:5F"


class BaseBanc(unittest.TestCase):
    def setUp(self):
        self.dossier = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.dossier, ignore_errors=True)

    def fiche(self, numero, **kwargs):
        chemin = os.path.join(self.dossier, "X301523_Fiche-de-Test N%s.xlsx" % numero)
        return fixtures.ecrire_xlsx(chemin, **kwargs)

    def options(self, simulation=False, **kwargs):
        return runner.Options(dossier=self.dossier, feuille=FEUILLE, cellule=CELLULE,
                              simulation=simulation, **kwargs)

    def cellule(self, numero):
        chemin = os.path.join(self.dossier, "X301523_Fiche-de-Test N%s.xlsx" % numero)
        return xlsxcell.lire_cellule(chemin, FEUILLE, CELLULE)

    @staticmethod
    def liste(*couples):
        return [releve.Releve(n, m, "192.168.0.100", "2026-09-11 10:14:32") for n, m in couples]

    @staticmethod
    def statuts(bilan):
        return sorted((r.numero, r.statut) for r in bilan.resultats)


class TestReleveVersExcel(BaseBanc):
    def test_analyse_puis_ecriture_puis_relance(self):
        self.fiche(N1)
        self.fiche(N2)
        liste = self.liste((N1, M1), (N2, M2))

        analyse = runner.executer_releves(self.options(simulation=True), liste)
        self.assertEqual(analyse.compter(runner.SIMULATION), 2)
        self.assertIsNone(self.cellule(N1), "l'analyse ne doit rien écrire")

        ecriture = runner.executer_releves(self.options(), liste)
        self.assertEqual(ecriture.compter(runner.ECRIT), 2)
        self.assertEqual(self.cellule(N1), M1)
        self.assertEqual(self.cellule(N2), M2)

        relance = runner.executer_releves(self.options(), liste)
        self.assertEqual(relance.compter(runner.DEJA_OK), 2)
        self.assertEqual(relance.compter(runner.ECRIT), 0)

    def test_numero_releve_sans_fiche(self):
        self.fiche(N1)
        bilan = runner.executer_releves(self.options(), self.liste((N1, M1), (N2, M2)))
        self.assertEqual(bilan.compter(runner.ECRIT), 1)
        ignores = [r for r in bilan.resultats if r.statut == runner.IGNORE]
        self.assertEqual([r.numero for r in ignores], [N2])
        self.assertIn("aucune fiche Excel", ignores[0].message)

    def test_fiche_sans_releve_reste_silencieuse(self):
        """Le banc écrit les appareils qu'on vient de passer, pas tout le dossier.

        Signaler les fiches non concernées noierait le journal : sur un dossier
        de 200 fiches dont 3 ont été relevées, les 197 autres ne sont pas des
        anomalies.
        """
        self.fiche(N1)
        self.fiche(N2)
        bilan = runner.executer_releves(self.options(), self.liste((N1, M1)))
        self.assertEqual(self.statuts(bilan), [(N1, runner.ECRIT)])
        self.assertIsNone(self.cellule(N2))

    def test_deux_mac_pour_un_meme_numero_dans_la_liste(self):
        """Une liste retouchée à la main peut se contredire : ne rien écrire alors."""
        self.fiche(N1)
        bilan = runner.executer_releves(self.options(), self.liste((N1, M1), (N1, M2)))
        self.assertEqual(bilan.compter(runner.ERREUR), 1)
        self.assertEqual(bilan.compter(runner.ECRIT), 0)
        self.assertIsNone(self.cellule(N1))

    def test_doublon_identique_traite_une_seule_fois(self):
        self.fiche(N1)
        bilan = runner.executer_releves(self.options(), self.liste((N1, M1), (N1, M1)))
        self.assertEqual(self.statuts(bilan), [(N1, runner.ECRIT)])

    def test_valeur_differente_protegee_sans_ecrasement(self):
        self.fiche(N1)
        runner.executer_releves(self.options(), self.liste((N1, M1)))
        bilan = runner.executer_releves(self.options(), self.liste((N1, M2)))
        self.assertEqual(bilan.compter(runner.IGNORE), 1)
        self.assertEqual(self.cellule(N1), M1, "la cellule ne doit pas avoir bougé")

    def test_ecrasement_explicite(self):
        self.fiche(N1)
        runner.executer_releves(self.options(), self.liste((N1, M1)))
        bilan = runner.executer_releves(self.options(ecraser=True), self.liste((N1, M2)))
        self.assertEqual(bilan.compter(runner.ECRIT), 1)
        self.assertEqual(self.cellule(N1), M2)

    def test_mac_invalide_dans_la_liste(self):
        self.fiche(N1)
        bilan = runner.executer_releves(self.options(), self.liste((N1, "FF:FF:FF:FF:FF:FF")))
        self.assertEqual(bilan.compter(runner.ERREUR), 1)
        self.assertIsNone(self.cellule(N1))

    def test_separateur_et_casse_configurables(self):
        self.fiche(N1)
        options = self.options(separateur_mac="-", mac_majuscules=False)
        runner.executer_releves(options, self.liste((N1, M1)))
        self.assertEqual(self.cellule(N1), "00-30-d6-4c-6e-05")

    def test_deux_fiches_pour_un_meme_numero(self):
        self.fiche(N1)
        fixtures.ecrire_xlsx(os.path.join(self.dossier, "doublon N%s.xlsx" % N1))
        bilan = runner.executer_releves(self.options(), self.liste((N1, M1)))
        self.assertEqual(bilan.compter(runner.ERREUR), 1)
        self.assertEqual(bilan.compter(runner.ECRIT), 0)

    def test_sous_dossiers(self):
        interne = os.path.join(self.dossier, "lot 1")
        os.makedirs(interne)
        fixtures.ecrire_xlsx(os.path.join(interne, "Fiche N%s.xlsx" % N1))
        liste = self.liste((N1, M1))
        self.assertEqual(runner.executer_releves(self.options(), liste).compter(runner.IGNORE), 1)
        bilan = runner.executer_releves(self.options(sous_dossiers=True), liste)
        self.assertEqual(bilan.compter(runner.ECRIT), 1)

    def test_dossier_introuvable(self):
        options = self.options()
        options.dossier = os.path.join(self.dossier, "absent")
        with self.assertRaises(ValueError):
            runner.executer_releves(options, self.liste((N1, M1)))

    def test_xls_refuse(self):
        self.fiche(N1)
        with open(os.path.join(self.dossier, "ancienne N%s.xls" % N2), "wb") as fichier:
            fichier.write(b"\xd0\xcf\x11\xe0")
        bilan = runner.executer_releves(self.options(), self.liste((N1, M1)))
        self.assertTrue(any(".xls" in r.message for r in bilan.resultats
                            if r.statut == runner.ERREUR))

    def test_interruption(self):
        self.fiche(N1)
        self.fiche(N2)
        bilan = runner.executer_releves(self.options(), self.liste((N1, M1), (N2, M2)),
                                        interruption=lambda: True)
        self.assertEqual(bilan.resultats, [])


class TestLesDeuxSourcesEcriventPareil(BaseBanc):
    """Un même appareil doit donner la même cellule, qu'il vienne du Word ou du banc.

    C'est la garantie qu'on peut passer d'une source à l'autre — ou mélanger les
    deux sur un même dossier — sans surprise.
    """

    def test_meme_resultat_depuis_le_word_et_depuis_le_banc(self):
        self.fiche(N1)
        fixtures.ecrire_docx(os.path.join(self.dossier, "X130392_B_%s.doc" % N1),
                             ["Fiche produit", "Adresse MAC : %s" % M1])

        runner.executer(self.options())
        depuis_word = self.cellule(N1)

        self.fiche(N1)                      # on repart d'une fiche vierge
        os.remove(os.path.join(self.dossier, "X130392_B_%s.doc" % N1))
        runner.executer_releves(self.options(), self.liste((N1, M1)))
        depuis_banc = self.cellule(N1)

        self.assertEqual(depuis_word, M1)
        self.assertEqual(depuis_banc, depuis_word)

    def test_le_chemin_word_ignore_la_liste(self):
        """La source Word ne doit rien devoir au CSV : son comportement est intact."""
        self.fiche(N1)
        fixtures.ecrire_docx(os.path.join(self.dossier, "X130392_B_%s.doc" % N1),
                             ["Adresse MAC : %s" % M1])
        bilan = runner.executer(self.options())
        self.assertEqual(self.statuts(bilan), [(N1, runner.ECRIT)])
        self.assertEqual(self.cellule(N1), M1)


if __name__ == "__main__":
    unittest.main()
