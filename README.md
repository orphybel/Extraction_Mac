# Extraction MAC — relever les adresses MAC, et les porter sur les fiches de test

Un seul programme pour les deux moitiés du travail :

1. **Relevé au banc** — les appareils sont branchés un par un sur le même accès
   LAN, où Tftpd32 leur donne toujours la même adresse IP. Le programme
   surveille cette adresse, détecte chaque appareil, et n'attend de l'opérateur
   que **les cinq derniers chiffres** du numéro de série.
2. **Écriture dans Excel** — chaque adresse MAC est portée dans la cellule
   voulue de la fiche de test correspondante, au format `xx:xx:xx:xx:xx:xx`.

Les adresses peuvent venir **du banc** ou, comme auparavant, **des fichiers
Word** fournis avec les appareils. Le reste du traitement est le même.

```
appareil branché ──┐
                   ├──► fiche de test .xlsx, cellule F27
fichier Word ──────┘
```

Les deux fichiers d'une même série sont appariés par le numéro présent dans
leur nom, de la forme `xxxxxx-yyyyyyy-zzzzz` :

```
X130392_B_260918-0144215-00092.doc                                          ─┐
X301523-9_MF19-Ecran-Cabine-12.1_Fiche-de-Test N°260918-0144215-00092.xlsx  ─┘  même série
```

Le reste du nom peut être quelconque, et les tirets du numéro sont facultatifs
(`260918014421500092` est reconnu de la même façon).

> **Vous veniez de la version précédente ?** Rien à ressaisir : le programme
> garde le même nom, le même exécutable et le même fichier de configuration.
> Vos profils — dossier, onglet, cellule, options — se rechargent tels quels, et
> l'écriture depuis les fichiers Word n'a pas changé d'un iota.

## Installation

### Option 1 — l'exécutable (rien à installer)

`ExtractionMAC.exe` est autonome : ni Python, ni bibliothèque, ni droits
administrateur pour le lancer. Le copier où l'on veut et double-cliquer.

Il est reconstruit à chaque modification du code par
[l'action GitHub « Construire l'exécutable Windows »](../../actions/workflows/build-exe.yml) :
ouvrir la dernière exécution réussie et télécharger l'artefact
**ExtractionMAC-windows**. Il contient l'exécutable et son empreinte SHA-256.

> Le premier lancement peut déclencher un avertissement SmartScreen
> (« Windows a protégé votre ordinateur ») : c'est le comportement normal pour
> un exécutable non signé. *Informations complémentaires* → *Exécuter quand même*.
> Certains antivirus signalent aussi à tort les exécutables PyInstaller ; comparer
> l'empreinte SHA-256 fournie permet de vérifier que le fichier est bien celui produit.

### Option 2 — depuis les sources

Le programme n'utilise que la bibliothèque standard de Python : **aucun
`pip install` n'est nécessaire**.

1. Installer Python 3.8 ou plus récent depuis <https://www.python.org/downloads/windows/>
   en cochant **« Add python.exe to PATH »** et **« tcl/tk and IDLE »**.
2. Double-cliquer sur `Lancer_ExtractionMac.bat`.

`Creer_executable.bat` refabrique l'exécutable localement, sur un poste Windows
disposant de Python.

## Utilisation

La fenêtre porte les **profils** en haut — ils mémorisent tous les réglages des
deux onglets — et deux onglets en dessous.

### Onglet « Relevé au banc »

| Champ | Rôle |
| --- | --- |
| **Adresse IP surveillée** | celle que Tftpd32 distribue, par exemple `192.168.0.100`. Le bouton *Tester* dit immédiatement qui répond. |
| **Préfixe du numéro** | le début du numéro de série, commun à toute la série : `260918-0144215`. Le bouton *Déduire du dossier* le trouve tout seul à partir des noms des fiches Excel. |
| **Fichier liste (CSV)** | le **chemin complet du fichier** où s'écrit la liste — pas un dossier. Il peut être posé n'importe où. |

> **Reprendre une liste déjà commencée.** Il suffit de **renseigner le dossier
> des fiches** : s'il contient un `MAC-releves.csv`, le champ se remplit tout
> seul et la liste est rechargée, prête à être continuée. Sinon, *Parcourir…*
> permet de désigner le fichier où qu'il soit. Le bouton *Recharger la liste*
> relit ce chemin et **dit toujours ce qu'il a trouvé** — y compris « aucun
> fichier choisi » ou « fichier introuvable ».
>
> Les listes écrites par la version précédente se rechargent telles quelles :
> le format n'a pas changé. Un dossier saisi par erreur à la place du fichier
> est toléré : la liste y est cherchée.

Puis, appareil par appareil :

1. **Démarrer la surveillance**.
2. Brancher l'appareil. Le programme sonne et affiche sa MAC en grand ; le
   curseur est déjà dans le champ du numéro.
3. Taper les derniers chiffres — `92` suffit pour `00092` — puis **Entrée**.
   Le numéro complet reconstitué s'affiche avant validation.
4. Débrancher l'appareil. Le programme vide le cache ARP et se réarme seul.
   Retour à l'étape 2.
5. En fin de série : **Écrire ces relevés dans les fiches Excel →**, qui bascule
   sur l'autre onglet avec la bonne source déjà choisie.

**Contrôles à la saisie.** Rien n'est ajouté à la liste en silence. Le programme
demande confirmation si le numéro est déjà dans la liste, si la MAC est déjà
relevée sous un autre numéro, ou si **aucune fiche Excel du dossier ne porte ce
numéro** — c'est ce qui attrape les fautes de frappe sur les cinq chiffres.

### Onglet « Écriture dans Excel »

1. **Dossier des fichiers** — celui qui contient les `.xlsx` (et les `.doc`, si
   les MAC en viennent). Ils peuvent être nombreux : tout est traité en une fois.
2. **Feuille Excel** — le bouton *Lire les onglets* remplit la liste à partir du
   premier classeur trouvé dans le dossier.
3. **Cellule** — par exemple `F27`. Le bouton *Trouver @MAC* liste les cellules
   contenant « MAC » dans le classeur, pour repérer où se trouve le libellé.
4. **Source des adresses MAC** — *fichiers Word du dossier* (comportement
   d'origine) ou *liste relevée au banc*.
5. **Analyser (sans écrire)** — vérifie tout sans modifier le moindre fichier.
   À faire systématiquement avant la première écriture sur un nouveau modèle de fiche.
6. **Écrire dans Excel** — demande confirmation, puis écrit.

Le **dossier** et le **fichier de la liste** sont partagés par les deux onglets :
les renseigner d'un côté les renseigne de l'autre.

#### Options

| Option | Effet |
| --- | --- |
| Inclure les sous-dossiers | parcourt aussi les dossiers contenus dans le dossier choisi |
| Écraser une valeur déjà présente | sans cette case, une cellule déjà remplie avec une **autre** valeur n'est jamais modifiée ; l'écart est signalé dans le journal |

Relancer le traitement sur un dossier déjà traité est sans risque : les
cellules déjà correctes sont signalées « DÉJÀ OK » et laissées telles quelles.

## Le point délicat : le cache ARP

C'est la seule vraie difficulté du relevé au banc, et elle mérite d'être comprise.

Windows garde en mémoire l'association « adresse IP → adresse MAC » pendant
quelques dizaines de secondes. Comme **tous les appareils reçoivent la même
IP**, après un échange rapide cette mémoire peut encore contenir la MAC de
l'appareil précédent. Sans précaution, le programme enregistrerait une MAC
fausse sous un numéro juste — une erreur silencieuse, donc coûteuse.

Trois garde-fous, dans cet ordre :

1. **Une MAC identique à celle qui vient d'être enregistrée n'est jamais
   proposée.** Le programme affiche « cache ARP non vidé ou appareil non
   remplacé » et attend.
2. **L'appareil doit être vu disparaître** avant que la détection se réarme. Un
   appareil laissé branché ne peut pas être compté deux fois.
3. **Le cache est vidé automatiquement** (`arp -d`) entre deux appareils. Cette
   commande demande les **droits administrateur**.

Si le vidage échoue, le programme le dit en barre d'état et propose un bouton
*Relancer en administrateur*. Sans élévation le programme reste utilisable, mais
il ne repose alors que sur les deux premiers garde-fous : **prendre le temps de
débrancher franchement chaque appareil**.

L'exécutable ne force pas l'élévation à chaque lancement, volontairement : cela
le rendrait inutilisable sur un poste sans droits administrateur.

## Ce que fait le programme, précisément

**Détection au banc.** L'adresse IP surveillée est sollicitée par un simple
datagramme UDP — suffisant pour provoquer la résolution ARP, sans dépendre du
ping qu'un appareil embarqué peut filtrer — puis la table ARP est lue. La
lecture ne dépend **ni de la langue de Windows ni des en-têtes** : sur chaque
ligne, on ne retient qu'une adresse IP et une adresse MAC. Une adresse de
diffusion, nulle ou de multidiffusion est refusée. Une MAC doit être vue
plusieurs fois d'affilée avant d'être déclarée détectée : pendant le
branchement, la table passe par des états transitoires.

**Lecture du Word.** Trois formats sont acceptés sans configuration :
`.docx`, `.doc` binaire Word 97-2003, et `.rtf` renommé en `.doc`. La recherche
privilégie toujours une ligne portant le libellé « MAC », et accepte les
notations `00-30-D6-4C-6E-05`, `00:30:D6:4C:6E:05`, `0030.D64C.6E05` et
`0030D64C6E05`. Un numéro de série ne peut pas être confondu avec une adresse
MAC. Si le document contient deux adresses MAC différentes, le fichier est
signalé en erreur plutôt que traité au hasard. La liste relevée au banc est
soumise au même contrôle : deux MAC pour un même numéro, et rien n'est écrit.

**Écriture de l'Excel.** Seule la feuille visée est réécrite ; toutes les autres
pièces du classeur sont recopiées **octet pour octet**. C'est un choix
délibéré : ces fiches contiennent des images, des contrôles de formulaire, des
réglages d'impression et des mises en forme conditionnelles qu'une réécriture
par une bibliothèque tierce ferait disparaître. Le style de la cellule
(bordures, police, fond) est conservé. La valeur est écrite en texte, donc
jamais réinterprétée par Excel.

Si la cellule visée appartient à une plage fusionnée, la valeur est écrite dans
la cellule d'ancrage de la fusion.

## Cas signalés dans le journal

| Statut | Signification |
| --- | --- |
| `ÉCRIT` | adresse MAC écrite dans la cellule |
| `À ÉCRIRE` | résultat d'une analyse : ce qui serait écrit |
| `DÉJÀ OK` | la cellule contient déjà cette adresse MAC |
| `IGNORÉ` | fichier Word sans Excel (ou l'inverse), relevé sans fiche, ou cellule déjà remplie avec une autre valeur |
| `ERREUR` | MAC introuvable ou ambiguë, onglet absent, fichier verrouillé, doublon de numéro… |

Une différence assumée entre les deux sources : **depuis le banc, une fiche
Excel sans relevé n'est pas signalée**. Le chemin Word traite tout un dossier ;
le chemin banc écrit les quelques appareils qu'on vient de passer. Signaler les
200 fiches non concernées noierait le journal.

Un fichier `.xls` (Excel 97-2003) ne peut pas être traité : le signaler et
l'enregistrer au format `.xlsx`. Un classeur ouvert dans Excel est verrouillé :
le fermer avant de relancer.

## Répétition à blanc, sans appareil

```
ExtractionMAC.exe --simulation
```

Des appareils fictifs se présentent l'un après l'autre : tout le déroulé du banc
fonctionne — détection, saisie, retrait, réarmement — et une vraie liste est
produite. Utile pour former un opérateur, ou pour vérifier un préfixe et un
dossier de fiches sans mobiliser le banc.

## Où sont enregistrés les profils

Un profil retient **les réglages des deux onglets** : dossier, onglet et cellule
où écrire, source des MAC, adresse IP surveillée, préfixe, fichier de la liste.

Ils sont enregistrés dans **`ExtractionMAC-config.json`, à côté de
l'exécutable** (ou à la racine du projet si l'on lance les sources). L'outil est
donc portable : copier le dossier — ou le mettre sur une clé USB — et les profils
suivent. La barre d'état affiche l'emplacement exact au démarrage.

> **Repli automatique.** Si ce dossier n'est pas accessible en écriture —
> exécutable posé dans `C:\Program Files`, sur un partage réseau en lecture
> seule, sur une clé protégée — la configuration bascule vers
> `%APPDATA%\ExtractionMac\`. Sans ce repli, l'enregistrement échouerait sans
> raison apparente. À la lecture, le fichier situé à côté de l'exécutable est
> prioritaire.
>
> En cas de doute, `ExtractionMAC.exe --diagnostic` écrit un fichier
> `ExtractionMAC-diagnostic.txt` indiquant les emplacements retenus.

## Développement

```
extraction_mac/
├── serial.py        le numéro de série : lecture dans un nom, composition au banc
├── docmac.py        lecture de l'adresse MAC dans le Word (docx, doc binaire, rtf)
├── arp.py           table ARP, sollicitation de l'IP, vidage du cache
├── surveillance.py  machine à états du banc : attente / détecté / attente du retrait
├── releve.py        la liste des relevés et son fichier CSV
├── simulateur.py    fausse table ARP, pour la répétition à blanc
├── xlsxcell.py      lecture et écriture chirurgicale d'une cellule .xlsx / .xlsm
├── runner.py        appairage et exécution, pour les deux sources de MAC
├── config.py        profils enregistrés (à côté du programme, repli %APPDATA%)
├── gui.py           fenêtre, profils, onglets
├── gui_banc.py      onglet « Relevé au banc »
└── gui_excel.py     onglet « Écriture dans Excel »
```

Les deux sources d'adresses partagent tout sauf une ligne : `runner` porte la
boucle commune (lecture de la cellule, `DÉJÀ OK`, protection contre
l'écrasement, simulation, écriture), et la source n'intervient que par la
fonction qui fournit la MAC.

Tests (144 cas, sans aucune donnée client, sans réseau et sans appareil — les
fichiers d'essai sont fabriqués à la volée) :

```
python -m unittest discover -s tests -t .
```

| Fichier | Ce qu'il garde |
| --- | --- |
| `test_extraction.py` | le comportement d'origine : lecture du Word, écriture chirurgicale, appairage, profils |
| `test_banc.py` | table ARP, composition du numéro, machine à états, liste CSV, simulateur |
| `test_source_releve.py` | la chaîne relevé → cellule Excel relue, et l'égalité des deux sources |
| `test_interface.py` | la fenêtre elle-même : elle se construit, les onglets travaillent, et aucune méthode ne masque un membre de Tkinter |

Les tests de fenêtre ont besoin d'un affichage ; ils sont sautés proprement s'il
n'y en a pas. Pour les exécuter sur un poste sans écran :

```
xvfb-run -a python -m unittest discover -s tests -t .
```

`ExtractionMAC.exe --autotest` construit la fenêtre, la referme et rend un code
de sortie : `0` si tout va bien, `1` en cas d'erreur — la trace est alors écrite
dans `ExtractionMAC-autotest.txt` —, `2` si aucun affichage n'est disponible.
C'est ce que vérifie l'action GitHub avant de publier l'exécutable.

## Réglages avancés

Certains réglages ne sont pas exposés dans la fenêtre mais modifiables dans
`ExtractionMAC-config.json`, par profil :

| Clé | Défaut | Rôle |
| --- | --- | --- |
| `separateur_mac` | `":"` | séparateur entre octets de l'adresse MAC |
| `mac_majuscules` | `true` | `false` pour écrire `00:30:d6:4c:6e:05` |
| `groupes_numero` | `[6, 7, 5]` | longueurs des trois groupes du numéro de série |
| `intervalle_ms` | `1500` | délai entre deux lectures de la table ARP |
| `scrutations_stables` | `3` | lectures identiques exigées avant de déclarer un appareil détecté |
| `scrutations_absence` | `2` | lectures à vide exigées avant de déclarer l'appareil retiré |
