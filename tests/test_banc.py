"""Tests du relevé au banc. Aucune donnée client, aucun réseau, aucun appareil."""

from __future__ import annotations

import os
import shutil
import tempfile
import unittest

from extraction_mac import arp, config, releve, runner, serial, simulateur, surveillance

A = "00:30:D6:4C:6E:05"
B = "00:30:D6:4C:6E:12"


# ==========================================================================
# lecture de la table ARP
# ==========================================================================

TABLE_FR = """
Interface : 192.168.0.10 --- 0xb
  Adresse Internet      Adresse physique      Type
  192.168.0.100         00-30-d6-4c-6e-05     dynamique
  192.168.0.255         ff-ff-ff-ff-ff-ff     statique
  224.0.0.22            01-00-5e-00-00-16     statique
"""

TABLE_EN = """
Interface: 192.168.0.10 --- 0xb
  Internet Address      Physical Address      Type
  192.168.0.100         00-30-d6-4c-6e-05     dynamic
"""


class TestArp(unittest.TestCase):
    def test_table_francaise(self):
        self.assertEqual(arp.analyser_table(TABLE_FR)["192.168.0.100"], A)

    def test_table_anglaise_donne_le_meme_resultat(self):
        """La lecture ne doit dépendre ni de la langue ni des en-têtes."""
        self.assertEqual(arp.analyser_table(TABLE_EN)["192.168.0.100"],
                         arp.analyser_table(TABLE_FR)["192.168.0.100"])

    def test_en_tete_sans_mac_est_ignore(self):
        self.assertNotIn("192.168.0.10", arp.analyser_table(TABLE_FR))

    def test_format_linux(self):
        self.assertEqual(arp.analyser_table("? (192.168.0.100) at 00:30:d6:4c:6e:05 [ether] on eth0"),
                         {"192.168.0.100": A})

    def test_format_ip_neigh(self):
        self.assertEqual(arp.analyser_table("192.168.0.100 dev eth0 lladdr 00:30:d6:4c:6e:05 STALE"),
                         {"192.168.0.100": A})

    def test_sortie_vide_ou_parasite(self):
        self.assertEqual(arp.analyser_table(""), {})
        self.assertEqual(arp.analyser_table("Aucune entree ARP trouvee"), {})
        self.assertEqual(arp.analyser_table(None), {})

    def test_ligne_a_deux_mac_est_ecartee(self):
        """Une ligne ambiguë vaut mieux ignorée que devinée."""
        self.assertEqual(
            arp.analyser_table("192.168.0.100 00-30-d6-4c-6e-05 00-30-d6-4c-6e-06"), {})

    def test_mise_en_forme_partagee_avec_docmac(self):
        """Une MAC relevée et une MAC lue dans un Word doivent s'écrire pareil."""
        from extraction_mac import docmac
        self.assertIs(arp.normaliser, docmac.normaliser)
        for brut in ("00-30-D6-4C-6E-05", "00:30:d6:4c:6e:05", "0030.D64C.6E05"):
            self.assertEqual(arp.normaliser(brut), A)

    def test_mac_de_diffusion_refusee(self):
        self.assertFalse(arp.mac_exploitable("FF:FF:FF:FF:FF:FF")[0])
        self.assertFalse(arp.mac_exploitable("00:00:00:00:00:00")[0])

    def test_mac_de_multidiffusion_refusee(self):
        """Premier octet impair : ce n'est jamais la MAC d'un appareil."""
        self.assertFalse(arp.mac_exploitable("01:00:5E:00:00:16")[0])

    def test_mac_incomplete_refusee(self):
        self.assertFalse(arp.mac_exploitable("00:30:D6")[0])

    def test_mac_normale_acceptee(self):
        self.assertTrue(arp.mac_exploitable(A)[0])


class TestVidageCacheArp(unittest.TestCase):
    """Le vidage ne doit pas crier au manque de droits quand il n'y a rien à vider."""

    def setUp(self):
        self.vraie_table = arp.table

    def tearDown(self):
        arp.table = self.vraie_table

    def test_entree_absente_n_est_pas_un_echec(self):
        arp.table = lambda: {}
        succes, message = arp.vider("192.168.0.100")
        self.assertTrue(succes)
        self.assertIn("aucune entrée", message)

    def test_entree_presente_declenche_la_commande(self):
        arp.table = lambda: {"192.168.0.100": A}
        appels = []
        vrai = arp._executer
        arp._executer = lambda arguments, delai=3: (appels.append(arguments), (1, "refusé"))[1]
        try:
            succes, message = arp.vider("192.168.0.100")
        finally:
            arp._executer = vrai
        self.assertEqual(appels, [["arp", "-d", "192.168.0.100"]])
        self.assertFalse(succes)
        self.assertIn("administrateur", message)


# ==========================================================================
# composition du numéro à partir des derniers chiffres
# ==========================================================================

class TestComposerNumero(unittest.TestCase):
    PREFIXE = "260918-0144215"

    def test_saisie_complete(self):
        self.assertEqual(serial.composer(self.PREFIXE, "00092")[0], "260918-0144215-00092")

    def test_zeros_ajoutes_a_gauche(self):
        """Le gain de temps recherché : taper 92 doit suffire."""
        self.assertEqual(serial.composer(self.PREFIXE, "92")[0], "260918-0144215-00092")
        self.assertEqual(serial.composer(self.PREFIXE, "7")[0], "260918-0144215-00007")

    def test_saisie_trop_longue_refusee(self):
        complet, erreur = serial.composer(self.PREFIXE, "123456")
        self.assertIsNone(complet)
        self.assertIn("trop long", erreur)

    def test_saisie_non_numerique_refusee(self):
        self.assertIn("chiffres", serial.composer(self.PREFIXE, "9a")[1])

    def test_saisie_vide_refusee(self):
        self.assertIsNone(serial.composer(self.PREFIXE, "   ")[0])

    def test_numero_complet_colle_accepte(self):
        """Coller un numéro entier ne doit pas bloquer l'opérateur."""
        self.assertEqual(serial.composer(self.PREFIXE, "260918-0144215-00092")[0],
                         "260918-0144215-00092")

    def test_prefixe_sans_separateurs(self):
        self.assertEqual(serial.composer("2609180144215", "92")[0], "260918-0144215-00092")

    def test_prefixe_trop_court_refuse(self):
        self.assertIn("trop court", serial.composer("2609", "92")[1])

    def test_prefixe_vide_refuse(self):
        self.assertIsNone(serial.composer("", "92")[0])

    def test_prefixe_normalise(self):
        self.assertEqual(serial.normaliser_prefixe("260918 0144215")[0], "260918-0144215")
        self.assertEqual(serial.normaliser_prefixe("260918-0144215-00092")[0], "260918-0144215")

    def test_numero_compose_est_relisible_par_l_appairage(self):
        """Le numéro composé doit être exactement celui lu dans un nom de fichier."""
        complet, _ = serial.composer(self.PREFIXE, "92")
        lu, erreur = serial.key_from_filename("Fiche N°%s.xlsx" % complet)
        self.assertIsNone(erreur)
        self.assertEqual(lu, complet)

    def test_deduction_du_prefixe(self):
        prefixe, _ = serial.deduire_prefixe(
            ["Fiche N°260918-0144215-00092.xlsx", "Fiche N°260918-0144215-00093.xlsx"])
        self.assertEqual(prefixe, "260918-0144215")

    def test_deduction_refusee_si_prefixes_differents(self):
        """Mieux vaut ne rien proposer qu'une valeur fausse."""
        prefixe, message = serial.deduire_prefixe(
            ["N°260918-0144215-00092.xlsx", "N°260919-0144215-00093.xlsx"])
        self.assertIsNone(prefixe)
        self.assertIn("plusieurs préfixes", message)

    def test_deduction_sans_numero(self):
        self.assertIsNone(serial.deduire_prefixe(["fiche.xlsx"])[0])

    def test_groupes_configurables(self):
        self.assertEqual(serial.composer("1234", "5", groups=(4, 3))[0], "1234-005")


# ==========================================================================
# machine à états du banc
# ==========================================================================

class TestSurveillance(unittest.TestCase):
    def setUp(self):
        self.s = surveillance.Surveillance(scrutations_stables=3, scrutations_absence=2)
        self.s.demarrer()

    def _scruter(self, mac, fois=1):
        dernier = None
        for _ in range(fois):
            evenement = self.s.scruter(mac)
            if evenement:
                dernier = evenement
        return dernier

    def test_detection_apres_trois_scrutations_stables(self):
        self.assertIsNone(self._scruter(A))
        self.assertIsNone(self._scruter(A))
        self.assertEqual(self._scruter(A), (surveillance.DETECTION, A))

    def test_une_apparition_fugace_ne_declenche_rien(self):
        """Pendant le branchement la table ARP passe par des états transitoires."""
        self._scruter(A)
        self._scruter(None)
        self._scruter(A)
        self._scruter(B)
        self.assertEqual(self.s.etat, surveillance.ATTENTE)

    def test_scenario_complet_deux_appareils(self):
        self.assertEqual(self._scruter(A, 3), (surveillance.DETECTION, A))
        self.assertTrue(self.s.enregistrer())
        self.assertEqual(self._scruter(None, 2), (surveillance.RETRAIT, None))
        self.assertEqual(self._scruter(B, 3), (surveillance.DETECTION, B))

    def test_garde_fou_mac_identique(self):
        """Le piège central : même IP, cache ARP pas vidé, appareil pas changé."""
        self._scruter(A, 3)
        self.s.enregistrer()
        self._scruter(None, 2)
        self.assertEqual(self._scruter(A, 3), (surveillance.IDENTIQUE, A))
        self.assertEqual(self.s.etat, surveillance.ATTENTE)
        self.assertIsNone(self.s.mac_courante)

    def test_mac_identique_signalee_une_seule_fois(self):
        self._scruter(A, 3)
        self.s.enregistrer()
        self._scruter(None, 2)
        self.assertEqual(self._scruter(A, 3), (surveillance.IDENTIQUE, A))
        self.assertIsNone(self._scruter(A, 5))

    def test_appareil_non_debranche_reste_bloquant(self):
        """Tant qu'on n'a pas vu l'appareil partir, on ne réarme pas."""
        self._scruter(A, 3)
        self.s.enregistrer()
        self.assertIsNone(self._scruter(A, 10))
        self.assertEqual(self.s.etat, surveillance.ATTENTE_RETRAIT)

    def test_echange_rapide_sans_absence_visible(self):
        self._scruter(A, 3)
        self.s.enregistrer()
        self.assertEqual(self._scruter(B), (surveillance.RETRAIT, None))
        self.assertEqual(self._scruter(B, 3), (surveillance.DETECTION, B))

    def test_appareil_perdu_avant_la_saisie(self):
        self._scruter(A, 3)
        self.assertEqual(self._scruter(None, 2), (surveillance.PERDU, None))

    def test_absence_breve_ne_perd_pas_l_appareil(self):
        self._scruter(A, 3)
        self.assertIsNone(self._scruter(None))
        self.assertIsNone(self._scruter(A))
        self.assertEqual(self.s.etat, surveillance.DETECTE)

    def test_autre_mac_pendant_la_saisie(self):
        self._scruter(A, 3)
        self._scruter(B)
        self.assertEqual(self.s.etat, surveillance.ATTENTE)
        self.assertEqual(self._scruter(B, 2), (surveillance.DETECTION, B))

    def test_ignorer_ne_redetecte_pas_le_meme(self):
        self._scruter(A, 3)
        self.assertTrue(self.s.ignorer())
        self._scruter(None, 2)
        self.assertEqual(self._scruter(A, 3), (surveillance.IDENTIQUE, A))

    def test_rearmement_manuel(self):
        self._scruter(A, 3)
        self.s.enregistrer()
        self.assertTrue(self.s.rearmer())
        self.assertEqual(self.s.etat, surveillance.ATTENTE)

    def test_arret_ne_produit_plus_d_evenement(self):
        self.s.arreter()
        self.assertIsNone(self._scruter(A, 10))

    def test_enregistrer_hors_detection_refuse(self):
        self.assertFalse(self.s.enregistrer())


# ==========================================================================
# la liste et son CSV
# ==========================================================================

class TestReleve(unittest.TestCase):
    def setUp(self):
        self.dossier = tempfile.mkdtemp()
        self.chemin = os.path.join(self.dossier, "liste.csv")

    def tearDown(self):
        shutil.rmtree(self.dossier, ignore_errors=True)

    def test_aller_retour(self):
        origine = [releve.Releve("260918-0144215-00092", A, "192.168.0.100", "2026-09-11 10:14:32"),
                   releve.Releve("260918-0144215-00093", B, "192.168.0.100", "2026-09-11 10:20:01")]
        releve.enregistrer(self.chemin, origine)
        relus, message = releve.charger(self.chemin)
        self.assertEqual([(r.numero, r.mac, r.ip, r.horodatage) for r in relus],
                         [(r.numero, r.mac, r.ip, r.horodatage) for r in origine])
        self.assertIn("2 relevé", message)

    def test_fichier_absent_n_est_pas_une_erreur(self):
        self.assertEqual(releve.charger(os.path.join(self.dossier, "rien.csv")), ([], ""))

    def test_separateur_point_virgule_pour_excel(self):
        releve.enregistrer(self.chemin, [releve.Releve("260918-0144215-00092", A)])
        with open(self.chemin, encoding="utf-8-sig") as fichier:
            self.assertEqual(fichier.readline().strip(), "numero;mac;ip;horodatage")

    def test_bom_present_pour_excel_francais(self):
        releve.enregistrer(self.chemin, [releve.Releve("260918-0144215-00092", A)])
        with open(self.chemin, "rb") as fichier:
            self.assertEqual(fichier.read(3), b"\xef\xbb\xbf")

    def test_relecture_d_un_csv_a_virgules(self):
        with open(self.chemin, "w", encoding="utf-8") as fichier:
            fichier.write("numero,mac,ip,horodatage\n260918-0144215-00092,%s,,\n" % A)
        self.assertEqual(releve.charger(self.chemin)[0][0].mac, A)

    def test_en_tete_inattendu_signale(self):
        with open(self.chemin, "w", encoding="utf-8") as fichier:
            fichier.write("a;b;c\n1;2;3\n")
        relus, message = releve.charger(self.chemin)
        self.assertEqual(relus, [])
        self.assertIn("en-tête", message)

    def test_lignes_incompletes_ignorees_et_signalees(self):
        with open(self.chemin, "w", encoding="utf-8") as fichier:
            fichier.write("numero;mac;ip;horodatage\n;%s;;\n260918-0144215-00092;%s;;\n" % (A, A))
        relus, message = releve.charger(self.chemin)
        self.assertEqual(len(relus), 1)
        self.assertIn("incomplète", message)

    def test_doublon_de_numero_signale(self):
        anomalies = releve.controler([releve.Releve("260918-0144215-00092", A)],
                                     "260918-0144215-00092", B)
        self.assertTrue(any("déjà dans la liste" in a for a in anomalies))

    def test_mac_deja_relevee_sous_un_autre_numero(self):
        anomalies = releve.controler([releve.Releve("260918-0144215-00092", A)],
                                     "260918-0144215-00093", A)
        self.assertTrue(any("déjà relevée" in a for a in anomalies))

    def test_aucune_anomalie_sur_un_cas_normal(self):
        self.assertEqual(releve.controler([releve.Releve("260918-0144215-00092", A)],
                                          "260918-0144215-00093", B), [])

    def test_ecriture_ne_laisse_pas_de_fichier_temporaire(self):
        releve.enregistrer(self.chemin, [releve.Releve("260918-0144215-00092", A)])
        self.assertEqual([n for n in os.listdir(self.dossier) if n.endswith(".tmp")], [])


# ==========================================================================
# réglages ajoutés au profil
# ==========================================================================

class TestProfilDuBanc(unittest.TestCase):
    def test_ancien_profil_se_recharge_sans_migration(self):
        """Le nom du fichier de configuration ne change pas : rien à ressaisir."""
        ancien = {"dossier": "D:/fiches", "feuille": "Constit produit", "cellule": "F27",
                  "sous_dossiers": True, "ecraser": False}
        profil = config.normaliser_profil(ancien)
        for cle, valeur in ancien.items():
            self.assertEqual(profil[cle], valeur)
        self.assertEqual(profil["source_mac"], "word")
        self.assertEqual(profil["ip_surveillee"], config.PROFIL_DEFAUT["ip_surveillee"])

    def test_valeurs_hors_bornes_ramenees(self):
        self.assertEqual(config.normaliser_profil({"intervalle_ms": 99999})["intervalle_ms"],
                         config.BORNES["intervalle_ms"][1])
        self.assertEqual(config.normaliser_profil({"scrutations_stables": 0})["scrutations_stables"],
                         1)

    def test_source_inconnue_ramenee_au_defaut(self):
        self.assertEqual(config.normaliser_profil({"source_mac": "n'importe quoi"})["source_mac"],
                         "word")

    def test_source_releve_acceptee(self):
        self.assertEqual(config.normaliser_profil({"source_mac": "releve"})["source_mac"], "releve")


# ==========================================================================
# répétition à blanc
# ==========================================================================

class TestSimulateur(unittest.TestCase):
    def test_l_appareil_reste_branche_jusqu_a_l_enregistrement(self):
        """Un opérateur lent ne doit pas voir l'appareil disparaître en pleine saisie."""
        faux = simulateur.SimulateurArp(["AA:AA:AA:AA:AA:AA", "BB:BB:BB:BB:BB:BB"])
        for _ in range(20):
            self.assertEqual(faux.lire("192.168.0.100"), "AA:AA:AA:AA:AA:AA")

    def test_sequence_complete(self):
        faux = simulateur.SimulateurArp(["AA:AA:AA:AA:AA:AA", "BB:BB:BB:BB:BB:BB"],
                                        scrutations_vides=2)
        self.assertEqual(faux.lire(), "AA:AA:AA:AA:AA:AA")
        faux.suivant()
        self.assertIsNone(faux.lire())
        self.assertIsNone(faux.lire())
        self.assertEqual(faux.lire(), "BB:BB:BB:BB:BB:BB")
        faux.suivant()
        self.assertTrue(faux.termine)

    def test_deroule_avec_la_machine_a_etats(self):
        """Répétition à blanc de bout en bout, sans appareil ni réseau."""
        faux = simulateur.SimulateurArp(["AA:AA:AA:AA:AA:AA", "BB:BB:BB:BB:BB:BB"],
                                        scrutations_vides=2)
        machine = surveillance.Surveillance(3, 2)
        machine.demarrer()
        detectees = []
        for _ in range(40):
            evenement = machine.scruter(faux.lire())
            if evenement and evenement[0] == surveillance.DETECTION:
                detectees.append(evenement[1])
                machine.enregistrer()
                faux.suivant()
        self.assertEqual(detectees, ["AA:AA:AA:AA:AA:AA", "BB:BB:BB:BB:BB:BB"])


# ==========================================================================
# repérage des fiches (contrôle d'une saisie)
# ==========================================================================

class TestNumerosDesFiches(unittest.TestCase):
    def setUp(self):
        self.dossier = tempfile.mkdtemp()

    def tearDown(self):
        shutil.rmtree(self.dossier, ignore_errors=True)

    def _creer(self, *noms):
        for nom in noms:
            chemin = os.path.join(self.dossier, nom)
            os.makedirs(os.path.dirname(chemin), exist_ok=True)
            open(chemin, "wb").close()

    def test_fiches_reperees(self):
        self._creer("Fiche N°260918-0144215-00092.xlsx", "Fiche N°260918-0144215-00093.xlsm")
        self.assertEqual(set(runner.numeros_des_fiches(self.dossier)),
                         {"260918-0144215-00092", "260918-0144215-00093"})

    def test_fichiers_temporaires_ignores(self):
        self._creer("~$Fiche N°260918-0144215-00092.xlsx")
        self.assertEqual(runner.numeros_des_fiches(self.dossier), {})

    def test_word_non_compte(self):
        self._creer("X130392_260918-0144215-00092.doc")
        self.assertEqual(runner.numeros_des_fiches(self.dossier), {})

    def test_sous_dossiers(self):
        self._creer(os.path.join("lot 1", "Fiche N°260918-0144215-00092.xlsx"))
        self.assertEqual(runner.numeros_des_fiches(self.dossier, False), {})
        self.assertEqual(set(runner.numeros_des_fiches(self.dossier, True)),
                         {"260918-0144215-00092"})

    def test_dossier_absent(self):
        self.assertEqual(runner.numeros_des_fiches("/dossier/qui/nexiste/pas"), {})


if __name__ == "__main__":
    unittest.main()
