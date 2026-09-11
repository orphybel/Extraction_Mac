"""Fausse table ARP, pour derouler tout le scenario sans appareil ni reseau.

Lance par ``EnregistrementMAC.exe --simulation``, le programme se comporte comme
au banc : des appareils fictifs se presentent l'un apres l'autre, on saisit leur
numero, on les voit disparaitre. Cela permet de former un operateur, de verifier
un prefixe et un dossier de fiches, et de produire un vrai CSV et de vrais
``.rtf`` — sans mobiliser le banc de test.

L'appareil fictif **reste branche tant qu'on ne l'a pas enregistre** : sinon un
operateur un peu lent verrait l'appareil disparaitre en pleine saisie, ce qui ne
represente pas le fonctionnement reel.
"""

from __future__ import annotations

MACS_FICTIVES = (
    "00:30:D6:4C:6E:05",
    "00:30:D6:4C:6E:12",
    "AC:DE:48:00:11:22",
    "0A:1B:2C:3D:4E:5F",
)


class SimulateurArp:
    """Remplace ``arp.lire`` : rend une MAC fictive, puis du vide, puis la suivante."""

    def __init__(self, macs=MACS_FICTIVES, scrutations_vides=3):
        self.macs = list(macs)
        self.scrutations_vides = max(1, int(scrutations_vides))
        self.index = 0
        self._vides_restantes = 0

    def lire(self, ip=None):
        """MAC de l'appareil fictif du moment, ou None pendant le « debranchement »."""
        if self._vides_restantes > 0:
            self._vides_restantes -= 1
            return None
        if self.index >= len(self.macs):
            return None
        return self.macs[self.index]

    def suivant(self):
        """Debranche l'appareil courant et prepare le suivant."""
        self.index += 1
        self._vides_restantes = self.scrutations_vides

    def vider(self, ip=None):
        return True, "cache ARP vidé (simulation)"

    @property
    def termine(self):
        return self.index >= len(self.macs)
