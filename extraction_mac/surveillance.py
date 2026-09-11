"""Machine a etats du releve : attendre, detecter, enregistrer, attendre le retrait.

Cette logique est volontairement separee de l'interface : elle ne connait ni
Tkinter ni le reseau, elle recoit seulement « la MAC vue a l'instant t » et dit
ce qui vient de se passer. C'est ce qui permet de la tester entierement sans
appareil et sans reseau.

Elle porte les deux garde-fous qui protegent du piege central du montage — tous
les appareils recoivent **la meme adresse IP**, et Windows garde l'association
IP/MAC en cache plusieurs dizaines de secondes :

1. une MAC identique a celle qui vient d'etre enregistree n'est **jamais**
   proposee a l'enregistrement : c'est le signe que le cache n'a pas ete vide ou
   que l'appareil n'a pas ete change ;
2. apres chaque enregistrement, on exige de **voir l'appareil disparaitre**
   avant de rearmer la detection.

Une MAC doit en outre etre vue plusieurs scrutations d'affilee avant d'etre
declaree detectee : pendant le branchement, la table ARP passe par des etats
transitoires qu'il ne faut pas prendre pour un appareil.
"""

from __future__ import annotations

ARRET = "arret"
ATTENTE = "attente"
DETECTE = "detecte"
ATTENTE_RETRAIT = "attente_retrait"

DETECTION = "detection"          # (DETECTION, mac)  nouvel appareil confirme
IDENTIQUE = "identique"          # (IDENTIQUE, mac)  meme MAC qu'au coup d'avant
RETRAIT = "retrait"              # (RETRAIT, None)   appareil debranche
PERDU = "perdu"                  # (PERDU, None)     disparu avant la saisie

LIBELLES = {
    ARRET: "Surveillance arrêtée.",
    ATTENTE: "En attente d'un appareil…",
    DETECTE: "Appareil détecté — saisir son numéro.",
    ATTENTE_RETRAIT: "Enregistré. Débrancher l'appareil pour passer au suivant.",
}


class Surveillance:
    """Suit l'etat du banc a partir des MAC successivement observees."""

    def __init__(self, scrutations_stables=3, scrutations_absence=2):
        self.scrutations_stables = max(1, int(scrutations_stables))
        self.scrutations_absence = max(1, int(scrutations_absence))
        self.etat = ARRET
        self.mac_courante = None            # MAC confirmee, en attente de saisie
        self.derniere_enregistree = None    # MAC du dernier appareil enregistre
        self._candidate = None
        self._compte = 0
        self._absences = 0
        self._identique_signalee = False

    # ------------------------------------------------------------- commandes
    def demarrer(self):
        self.etat = ATTENTE
        self._reinitialiser_attente()

    def arreter(self):
        self.etat = ARRET
        self.mac_courante = None
        self._reinitialiser_attente()

    def enregistrer(self):
        """A appeler quand l'operateur a valide le numero de l'appareil detecte."""
        if self.etat != DETECTE:
            return False
        self.derniere_enregistree = self.mac_courante
        self.mac_courante = None
        self.etat = ATTENTE_RETRAIT
        self._reinitialiser_attente()
        return True

    def ignorer(self):
        """Passer l'appareil detecte sans l'enregistrer.

        On retient tout de meme sa MAC : sans cela il serait immediatement
        redetecte, et l'operateur tournerait en rond sur le meme appareil.
        """
        if self.etat != DETECTE:
            return False
        self.derniere_enregistree = self.mac_courante
        self.mac_courante = None
        self.etat = ATTENTE_RETRAIT
        self._reinitialiser_attente()
        return True

    def rearmer(self):
        """Reprendre la detection sans attendre de voir l'appareil disparaitre."""
        if self.etat == ARRET:
            return False
        self.mac_courante = None
        self.etat = ATTENTE
        self._reinitialiser_attente()
        return True

    # -------------------------------------------------------------- interne
    def _reinitialiser_attente(self):
        self._candidate = None
        self._compte = 0
        self._absences = 0
        self._identique_signalee = False

    # ------------------------------------------------------------ scrutation
    def scruter(self, mac_vue):
        """Fait avancer la machine d'un pas. Retourne (evenement, mac) ou None."""
        if self.etat == ARRET:
            return None

        if self.etat == ATTENTE_RETRAIT:
            if mac_vue is None:
                self._absences += 1
                if self._absences >= self.scrutations_absence:
                    self.etat = ATTENTE
                    self._reinitialiser_attente()
                    return (RETRAIT, None)
                return None
            self._absences = 0
            if mac_vue == self.derniere_enregistree:
                return None
            # Appareil deja remplace sans qu'on ait vu l'absence : on rearme et
            # la detection du nouveau se fera a la scrutation suivante.
            self.etat = ATTENTE
            self._reinitialiser_attente()
            return (RETRAIT, None)

        if self.etat == DETECTE:
            if mac_vue is None:
                self._absences += 1
                if self._absences >= self.scrutations_absence:
                    self.etat = ATTENTE
                    self.mac_courante = None
                    self._reinitialiser_attente()
                    return (PERDU, None)
                return None
            self._absences = 0
            if mac_vue == self.mac_courante:
                return None
            # une autre MAC est apparue : on repart en attente et on la compte
            self.etat = ATTENTE
            self.mac_courante = None
            self._reinitialiser_attente()

        # etat ATTENTE
        if mac_vue is None:
            self._candidate = None
            self._compte = 0
            self._identique_signalee = False
            return None

        if self._candidate != mac_vue:
            self._candidate = mac_vue
            self._compte = 0
            self._identique_signalee = False
        self._compte += 1

        if self._compte < self.scrutations_stables:
            return None

        if mac_vue == self.derniere_enregistree:
            if self._identique_signalee:
                return None
            self._identique_signalee = True
            return (IDENTIQUE, mac_vue)

        self.etat = DETECTE
        self.mac_courante = mac_vue
        self._absences = 0
        return (DETECTION, mac_vue)
