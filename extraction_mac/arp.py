"""Lecture de l'adresse MAC de l'appareil branche, par la table ARP du poste.

Principe : Tftpd32 attribue toujours la **meme** adresse IP a l'appareil du
moment. Il suffit donc de solliciter cette IP puis de lire la table ARP pour
connaitre la MAC de l'appareil actuellement branche.

Trois precautions qui font toute la difference en production :

* **Aucune fenetre noire.** Le programme est fabrique en mode fenetre
  (``--windowed``) : sans ``CREATE_NO_WINDOW``, chaque appel a ``arp`` ferait
  clignoter une console. Comme on scrute en permanence, ce serait inutilisable.
* **Aucune dependance a la langue de Windows.** On ne lit ni les en-têtes ni le
  mot « dynamique »/« dynamic » : on repere sur chaque ligne une adresse IP et
  une adresse MAC, et on les apparie seulement si la ligne en contient
  exactement une de chaque. La meme lecture fonctionne sur la sortie de
  ``arp -a`` sous Windows (toutes langues), sous Linux, et sur ``ip neigh``.
* **Sollicitation sans ICMP.** Un simple datagramme UDP vers le port 9
  (« discard ») suffit a declencher la resolution ARP, sans lancer de processus
  et sans dependre du ping, qu'un appareil embarque peut filtrer.
"""

from __future__ import annotations

import os
import re
import socket
import subprocess

from . import docmac

_IP = re.compile(r"(?<![0-9.])((?:\d{1,3}\.){3}\d{1,3})(?![0-9.])")
_MAC = re.compile(r"(?<![0-9A-Za-z])([0-9A-Fa-f]{2}(?:[:-][0-9A-Fa-f]{2}){5})(?![0-9A-Za-z])")

PORT_SOLLICITATION = 9          # « discard », RFC 863

# La mise en forme des adresses est celle de docmac : une MAC relevee au banc et
# une MAC lue dans un Word doivent s'ecrire exactement pareil, sans quoi le cas
# « DEJA OK » ne serait pas reconnu d'une source a l'autre.
normaliser = docmac.normaliser


def mac_exploitable(mac):
    """Retourne (vrai_faux, raison). Ecarte la diffusion, le nul et le multicast."""
    hexa = re.sub(r"[^0-9A-Fa-f]", "", mac or "").upper()
    if len(hexa) != 12:
        return False, "adresse MAC incomplète"
    if hexa in docmac.MAC_INVALIDES:
        return False, "adresse MAC nulle ou de diffusion"
    if int(hexa[:2], 16) & 1:
        return False, "adresse MAC de multidiffusion (ce n'est pas un appareil)"
    return True, ""


def _sans_console():
    """Options de lancement qui evitent toute fenetre de console sous Windows."""
    if os.name != "nt":
        return {}
    demarrage = subprocess.STARTUPINFO()
    demarrage.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    demarrage.wShowWindow = 0                      # SW_HIDE
    return {
        "startupinfo": demarrage,
        "creationflags": getattr(subprocess, "CREATE_NO_WINDOW", 0x08000000),
    }


# Delai court : la scrutation tourne sur le fil de l'interface, et un
# « arp » qui trainerait s'y verrait comme un gel de la fenetre.
def _executer(arguments, delai=3):
    """Retourne (code_retour, sortie). Le code -1 signale une commande absente."""
    try:
        termine = subprocess.run(
            arguments,
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            timeout=delai,
            **_sans_console()
        )
    except (OSError, subprocess.SubprocessError) as erreur:
        return -1, str(erreur)
    # La console Windows n'est pas en UTF-8 : on decode sans jamais echouer.
    return termine.returncode, termine.stdout.decode("cp1252", "replace")


def analyser_table(sortie):
    """Retourne {ip: mac} a partir du texte d'une commande ``arp``/``ip neigh``.

    Une ligne n'est retenue que si elle porte exactement une IP et une MAC :
    les en-têtes (« Interface : 192.168.0.10 --- 0xb ») n'ont pas de MAC et
    sont ecartes d'eux-memes, quelle que soit la langue du poste.
    """
    table = {}
    for ligne in (sortie or "").splitlines():
        adresses_ip = _IP.findall(ligne)
        adresses_mac = _MAC.findall(ligne)
        if len(adresses_ip) == 1 and len(adresses_mac) == 1:
            table[adresses_ip[0]] = normaliser(adresses_mac[0])
    return table


def table():
    """Table ARP du poste, sous forme {ip: mac}."""
    code, sortie = _executer(["arp", "-a"])
    if code == 0:
        resultat = analyser_table(sortie)
        if resultat:
            return resultat
    if os.name != "nt":                            # poste de developpement
        code, sortie = _executer(["ip", "neigh", "show"])
        if code == 0:
            return analyser_table(sortie)
    return {}


def solliciter(ip):
    """Provoque la resolution ARP de *ip* sans lancer de processus."""
    try:
        prise = socket.socket(socket.AF_INET, socket.SOCK_DGRAM)
    except OSError:
        return
    try:
        prise.settimeout(0.2)
        prise.sendto(b"\x00", (ip, PORT_SOLLICITATION))
    except OSError:
        pass                                       # pas de reponse attendue
    finally:
        prise.close()


def lire(ip):
    """MAC actuellement associee a *ip*, ou None. Sollicite l'IP au passage."""
    solliciter(ip)
    mac = table().get(ip)
    if not mac:
        return None
    exploitable, _ = mac_exploitable(mac)
    return mac if exploitable else None


def vider(ip):
    """Efface l'entree ARP de *ip*. Retourne (succes, message).

    ``arp -d`` exige les droits administrateur sous Windows. L'echec n'est
    jamais avale en silence : sans vidage, la table peut encore contenir la MAC
    de l'appareil precedent, et c'est exactement l'erreur que le programme doit
    eviter.

    On regarde d'abord si l'entree existe. Sans cela, le vidage declenche au
    moment ou l'appareil vient d'etre debranche tomberait tres souvent sur une
    entree deja disparue : ``arp -d`` renverrait une erreur, et le programme
    reclamerait les droits administrateur sans aucune raison. Une fausse alerte
    repetee finit par etre ignoree — y compris le jour ou elle est fondee.
    """
    if ip not in table():
        return True, "aucune entrée ARP à vider pour %s" % ip
    code, sortie = _executer(["arp", "-d", ip])
    if code == 0:
        return True, "cache ARP vidé pour %s" % ip
    if code == -1:
        return False, "commande arp introuvable (%s)" % sortie.strip()
    detail = " ".join(sortie.split())[:160]
    return False, ("vidage du cache ARP refusé (droits administrateur requis)"
                   + (" : %s" % detail if detail else ""))


def elevation_possible():
    return os.name == "nt"


def relancer_en_administrateur():
    """Relance le programme avec elevation. Retourne (succes, message).

    On ne met volontairement pas de manifeste UAC sur l'executable : cela
    imposerait l'elevation a chaque lancement et bloquerait un poste sans droits
    administrateur. L'elevation reste donc un geste volontaire de l'operateur.
    """
    if os.name != "nt":
        return False, "élévation disponible uniquement sous Windows"
    import sys
    import ctypes
    if getattr(sys, "frozen", False):
        programme, parametres = sys.executable, ""
    else:
        programme = sys.executable
        parametres = " ".join('"%s"' % a for a in [os.path.abspath(sys.argv[0])] + sys.argv[1:])
    try:
        code = ctypes.windll.shell32.ShellExecuteW(None, "runas", programme,
                                                   parametres or None, None, 1)
    except (AttributeError, OSError) as erreur:
        return False, "élévation impossible : %s" % erreur
    if code <= 32:
        return False, "élévation refusée (code %d)" % code
    return True, "programme relancé en administrateur"
