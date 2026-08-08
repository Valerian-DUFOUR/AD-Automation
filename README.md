# AD Automation — Boîte à outils Active Directory & réseau

Collection de scripts **PowerShell** pour l'administration, l'hygiène et la
**sécurité** d'un domaine Active Directory : extractions (utilisateurs,
ordinateurs, groupes, GPO), détection de doublons, comptes et postes inactifs,
couverture BitLocker, posture de sécurité (type PingCastle allégé), audit des
partages SMB/NTFS et inventaire réseau multi-sites. La plupart des scripts
génèrent un rapport **Excel (.xlsx)** avec une page **KPI** et des dates au
format **JJ-MM-AAAA**.

## Vue d'ensemble

| Script | Rôle | Sortie |
|---|---|---|
| `Export-ADUsersReport.ps1` | Comptes utilisateurs AD | Excel (KPI + détail) |
| `Export-ADComputersReport.ps1` | Ordinateurs AD + obsolescence Windows | Excel (KPI + détail) |
| `Export-ADGroupsReport.ps1` | Groupes AD et leurs membres (users / PC) | Excel (4 feuilles) |
| `Export-GPOReport.ps1` | Stratégies de groupe (GPO), liens, fichiers scripts | Excel (4 feuilles) |
| `Export-ADDuplicateObjects.ps1` | Objets en double (CNF, SPN, UPN, e-mail…) | Excel (KPI + détail) |
| `Export-ADSecurityPosture.ps1` | Posture de sécurité RSSI (type PingCastle) | Excel (KPI + 8 feuilles) |
| `Export-ADInactiveReport.ps1` | Comptes & PC inactifs (30 / 90 / 180 / 365 j) | Excel (KPI + détail) |
| `Export-ADComputersNoBitLocker.ps1` | Postes sans clé BitLocker dans l'AD | Excel (KPI + détail) |
| `Invoke-NetworkInventory.ps1` | Inventaire réseau + conformité VLAN | CSV / JSON / Excel (KPI) |
| `Audit_Securite_Serveurs_DC.ps1` + `Formater_Rapport_Excel.ps1` | Audit des partages SMB/NTFS (2 étapes) | CSV/JSON puis Excel |

**Fichiers de données fournis :** `oui-db.csv` (base des fabricants MAC IEEE,
requise par l'inventaire réseau), `sites.example.csv` et `vlans.example.csv`
(modèles de plan à adapter pour l'inventaire réseau).

Voir aussi la section [Prérequis](#prérequis) et le
[Guide pas à pas de l'audit SMB/NTFS](#guide--audit-des-partages-smbntfs-pas-à-pas).

## Scripts en détail

### `Export-ADUsersReport.ps1`

Extrait l'ensemble des comptes utilisateurs d'Active Directory et produit un
rapport Excel (`.xlsx`) déposé sur le Bureau.

Colonnes exportées :

| Colonne | Attribut AD |
|---|---|
| Nom du compte | `SamAccountName` |
| Email | `mail` |
| Prénom | `GivenName` |
| Nom | `Surname` |
| Nom complet | `DisplayName` |
| Description | `Description` |
| Dernière connexion | `LastLogonDate` |
| Dernier changement MDP | `PasswordLastSet` |
| Date de création | `whenCreated` |
| Emplacement (OU) | `DistinguishedName` |
| Statut (Activé / Désactivé) | `Enabled` |
| MDP n'expire jamais | `PasswordNeverExpires` |

Le classeur contient **deux feuilles** (dates au format JJ-MM-AAAA) :

1. **Synthèse (KPI)** — total de comptes, activés / désactivés, mots de passe
   « n'expire jamais », comptes sans e-mail, répartition par OU et graphique.
2. **Utilisateurs AD** — le détail en tableau filtrable : comptes **désactivés**
   surlignés en rouge clair, mots de passe **« n'expire jamais »** en rouge gras,
   colonnes ajustées et première ligne figée.

### `Export-ADComputersReport.ps1`

Extrait l'ensemble des ordinateurs d'Active Directory et produit un rapport
Excel (`.xlsx`) déposé sur le Bureau, avec une **page de synthèse (KPI)**.

Colonnes exportées :

| Colonne | Attribut AD |
|---|---|
| Nom du PC | `Name` |
| Description | `Description` |
| OS | `OperatingSystem` |
| Version OS | `OperatingSystemVersion` |
| Date de création | `whenCreated` |
| Dernière connexion | `LastLogonDate` |
| Statut (Activé / Désactivé) | `Enabled` |
| Jours inactif | calculé (aujourd'hui − `LastLogonDate`) |
| Emplacement (OU) | `DistinguishedName` |
| Obsolète | déduit de la version de Windows (fin de support) |
| Motif obsolescence | explication de l'obsolescence |

Le classeur contient **deux feuilles** :

1. **Synthèse** — indicateurs clés (total, activés/désactivés, obsolètes,
   inactifs), répartition par OS et graphique.
2. **Ordinateurs** — le détail de tous les postes, en tableau filtrable :
   postes **obsolètes** surlignés, **désactivés** en gris italique, et
   nombre de **jours d'inactivité** au-delà du seuil signalé en rouge.

Logique d'obsolescence (adaptable dans le script) : Windows 10 est considéré
obsolète (fin de support le 14/10/2025), Windows 11 et Server 2016+ sont
supportés, les versions antérieures sont marquées obsolètes.

Options :

```powershell
.\Export-ADComputersReport.ps1
.\Export-ADComputersReport.ps1 -SearchBase "OU=Postes,DC=example,DC=local" -InactiveDays 60
```

### `Export-ADGroupsReport.ps1`

Extrait tous les groupes AD et leurs membres, dans un classeur à quatre
feuilles (dates au format JJ-MM-AAAA) :

1. **Synthèse** — KPI : total de groupes, sécurité vs distribution, groupes
   vides, total des appartenances (dont utilisateurs / ordinateurs),
   répartition par étendue et graphique.
2. **Groupes** — un groupe par ligne : description, notes (« ce que fait le
   groupe »), catégorie, étendue, gestionnaire, date de création, emplacement
   (OU) et nombre de membres ventilé par type (utilisateurs / ordinateurs /
   groupes / autres). Les groupes vides sont surlignés.
3. **Membres Utilisateurs** — un couple (groupe, utilisateur) par ligne, avec
   les mêmes colonnes que `Export-ADUsersReport.ps1`.
4. **Membres Ordinateurs** — un couple (groupe, ordinateur) par ligne, avec
   les mêmes colonnes que `Export-ADComputersReport.ps1` (dont l'obsolescence).

```powershell
.\Export-ADGroupsReport.ps1
.\Export-ADGroupsReport.ps1 -SearchBase "OU=Groupes,DC=example,DC=local"
```

### `Export-GPOReport.ps1`

Extrait toutes les stratégies de groupe (module `GroupPolicy`), dans un
classeur à quatre feuilles (dates au format JJ-MM-AAAA) :

1. **Synthèse** — KPI : total de GPO, activées / désactivées, GPO non liées,
   GPO avec fichiers scripts, répartition par statut et graphique.
2. **GPO** — une GPO par ligne : statut, description, **date de création**,
   date de modification, propriétaire, filtre WMI, **chemin du dossier
   SYSVOL**, nombre de liens et nombre d'entités de filtrage de sécurité.
3. **Liens** — détail des liens (une ligne par OU/site où la GPO est appliquée).
4. **Fichiers scripts** — **chemins des fichiers** de scripts référencés par
   les GPO (ouverture/fermeture de session, démarrage/arrêt), s'il y en a.

> À propos du « nombre de membres par GPO » : une GPO n'a pas de membres au
> sens strict. Le rapport fournit deux mesures qui déterminent à qui la GPO
> s'applique : le **nombre de liens** (OU/sites) et le **nombre d'entités de
> filtrage de sécurité** (comptes/groupes ayant « Appliquer la stratégie de
> groupe »).

```powershell
.\Export-GPOReport.ps1
```

### `Export-ADDuplicateObjects.ps1`

Détecte les objets en double dans l'AD (hygiène et sécurité), dans un classeur
à deux feuilles (dates au format JJ-MM-AAAA) :

1. **Synthèse** — KPI par famille de doublon (valeurs en double et objets
   concernés), total et graphique.
2. **Doublons** — le détail, un objet concerné par ligne (type de doublon,
   valeur dupliquée, nombre d'occurrences, objet, OU, date de création).

Familles détectées :

| Famille | Intérêt |
|---|---|
| Objets en conflit (CNF) | Objets `...CNF:` issus de conflits de réplication |
| **SPN en double** | Même ServicePrincipalName sur plusieurs objets — casse Kerberos (critique) |
| UPN en double | `userPrincipalName` identiques |
| E-mail en double | `mail` / `proxyAddresses` en collision |
| DisplayName en double | Noms d'affichage identiques |
| Nom (CN) en double | Même nom dans des OU différentes |
| employeeID en double | Identifiants RH en double |

Les SPN en double (critiques) sont surlignés en rouge, les objets en conflit en
orange.

```powershell
.\Export-ADDuplicateObjects.ps1
.\Export-ADDuplicateObjects.ps1 -SearchBase "OU=Comptes,DC=example,DC=local"
```

### `Export-ADSecurityPosture.ps1`

Tableau de bord de **posture de sécurité** de l'AD (type PingCastle allégé),
orienté RSSI. **Lecture seule**, dates au format JJ-MM-AAAA.

- **Synthèse (KPI)** — première page : chaque indicateur avec un **niveau de
  risque** (OK / À surveiller / Critique, code couleur) et un graphique de
  répartition. Indicateurs : membres Domain / Enterprise / Schema Admins, total
  de comptes privilégiés, comptes privilégiés à mot de passe éternel ou
  inactifs, âge du mot de passe **krbtgt**, comptes et ordinateurs inactifs,
  mots de passe « n'expire jamais » / « non requis », comptes **Kerberoastables**,
  **délégation non contrainte**, objets avec **SIDHistory**, compte invité activé.
- Une **feuille de détail par thème** pour simplifier la lecture : `Comptes
  privilégiés`, `Comptes inactifs`, `Ordinateurs inactifs`, `Problèmes MDP`,
  `Kerberoastables`, `Délégation`, `SIDHistory`, `Politique MDP`.

```powershell
.\Export-ADSecurityPosture.ps1
.\Export-ADSecurityPosture.ps1 -InactiveDays 60 -MaxPasswordAgeDays 180
```

> Les seuils et niveaux de risque sont indicatifs et à adapter à votre
> contexte. Ce rapport ne remplace pas un audit complet (PingCastle,
> PurpleKnight…) mais fournit des KPI exploitables rapidement.

### `Export-ADInactiveReport.ps1`

Reprend les informations des rapports utilisateurs et ordinateurs, filtrées sur
les objets **inactifs (> 30 jours)**, dans un classeur à trois feuilles (dates
au format JJ-MM-AAAA) :

1. **Synthèse (KPI)** — nombre de comptes et d'ordinateurs inactifs de **plus de
   30, 90, 180 et 365 jours** (comptage cumulatif), avec un graphique comparatif
   utilisateurs / ordinateurs.
2. **Utilisateurs inactifs** — mêmes colonnes que `Export-ADUsersReport.ps1`,
   plus le nombre de jours d'inactivité et la tranche.
3. **Ordinateurs inactifs** — mêmes colonnes que `Export-ADComputersReport.ps1`,
   plus la tranche.

L'inactivité s'appuie sur `LastLogonDate` (ou la date de création si aucune
connexion n'est enregistrée, pour ne pas signaler à tort un objet récent).

```powershell
.\Export-ADInactiveReport.ps1
.\Export-ADInactiveReport.ps1 -SearchBase "OU=Comptes,DC=example,DC=local"
```

### `Export-ADComputersNoBitLocker.ps1`

Liste les postes **sans clé de récupération BitLocker sauvegardée dans l'AD**,
avec les mêmes colonnes que l'export des ordinateurs. Classeur à deux feuilles
(dates au format JJ-MM-AAAA) :

1. **Synthèse (KPI)** — total de postes, postes avec / sans clé, **taux de
   couverture** et graphique.
2. **PC sans BitLocker** — le détail des postes non couverts.

Détection : un poste dont la clé est sauvegardée dans l'AD possède un objet
enfant `msFVE-RecoveryInformation`. Le script recense ces objets et en déduit
les postes non couverts.

> Ceci vérifie la présence d'une clé **dans l'AD**, pas le chiffrement réel du
> disque (qui nécessite d'interroger la machine via `Get-BitLockerVolume`). La
> lecture des objets BitLocker requiert des droits d'administration.

```powershell
.\Export-ADComputersNoBitLocker.ps1
```

### `Invoke-NetworkInventory.ps1`

Inventaire réseau multi-sites et **contrôle de conformité VLAN**. Balaye les
plages IP (sites × VLAN) du plan d'adressage, récupère l'adresse MAC de chaque
équipement (ARP local, table ARP des passerelles via SNMP), résout le fabricant
via la base OUI IEEE (`oui-db.csv`), déduit le **type d'équipement** (PC,
serveur, NAS, imprimante, borne WiFi, switch, caméra IP, automate, onduleur…),
localise la machine (site + VLAN) et vérifie qu'elle est dans le **bon VLAN**.

Sorties : CSV, JSON, CSV des non-conformités, et Excel avec une feuille **KPI**
(total d'hôtes, conformes / tolérés / non conformes avec pourcentages, nombre de
sites et de types, répartition par site et par type + graphique) et une feuille
détaillée où les non-conformités sont surlignées.

Générique : **aucune donnée d'entreprise n'est codée en dur**. Plan d'adressage
au format **`192.A.B.C`** (A = 2ᵉ octet = site, B = 3ᵉ octet = VLAN, C = hôte),
le 1ᵉʳ octet étant réglable via `-BaseOctet`.

```powershell
# Nom d'entreprise + plans de sites ET de VLAN externes + moisson SNMP
.\Invoke-NetworkInventory.ps1 -Entreprise "ACME" -SitesFile .\sites.csv -VlansFile .\vlans.csv -UseSnmpGateways

# Cibler certains sites / VLAN, base d'adressage personnalisée
.\Invoke-NetworkInventory.ps1 -BaseOctet 10 -Sites 1,2 -Vlans 10,20,30 -UseSnmpGateways
```

Deux plans, **entièrement personnalisables** (rien de figé) :

- **Sites** (2ᵉ octet → établissement) : à éditer dans le script, ou via
  `-SitesFile` pointant un CSV `Octet;Societe;Ville;Id` (voir `sites.example.csv`).
- **VLAN** (3ᵉ octet → usage + types attendus) et **règles de conformité** par
  type d'équipement : **vides par défaut**, à remplir dans le script (blocs
  `$VlanMap` et `$TypeToVlan` clairement commentés) ou via `-VlansFile` pointant
  un CSV `Octet;Nom;Attendu` (voir `vlans.example.csv`). Sans plan VLAN, le
  balayage prend par défaut les VLAN 0 à 50 et la conformité est neutralisée.

À lancer de préférence **depuis le contrôleur de domaine** (ou un hôte routé vers
tous les sites) pour une couverture maximale. Nécessite `oui-db.csv` à côté du
script (base publique IEEE) ; l'export Excel utilise ImportExcel s'il est présent.

### `Audit_Securite_Serveurs_DC.ps1` + `Formater_Rapport_Excel.ps1`

Audit des partages réseau SMB/NTFS des serveurs Windows du domaine, en **deux
étapes** :

1. `Audit_Securite_Serveurs_DC.ps1` s'exécute **sur le contrôleur de domaine**
   (ou une machine avec RSAT). Il cible les serveurs de l'AD, teste la
   connectivité (Ping + WinRM), audite les droits SMB/NTFS (exposition
   « Everyone », contrôle total, propriétaire), classe chaque partage par
   niveau de risque, puis exporte un **CSV** et un **JSON de métadonnées**.
2. `Formater_Rapport_Excel.ps1` s'exécute **sur le PC de la personne**. Il lit
   le CSV + JSON et génère un **.xlsx** mis en forme (feuille de résumé avec
   graphique, feuille détaillée, feuilles par niveau de risque, serveurs
   injoignables). Nécessite Python 3 + openpyxl (installés automatiquement).

```powershell
# 1) Sur le contrôleur de domaine
.\Audit_Securite_Serveurs_DC.ps1
.\Audit_Securite_Serveurs_DC.ps1 -SearchBase "OU=Serveurs,DC=example,DC=local"

# 2) Sur le PC local, après avoir copié le CSV et le JSON
.\Formater_Rapport_Excel.ps1
```

> Ces deux scripts ne contiennent aucune donnée d'environnement : domaine,
> serveurs et utilisateurs sont détectés dynamiquement à l'exécution. Les
> exemples utilisent le domaine fictif `example.local`.

## Guide : audit des partages SMB/NTFS (pas à pas)

Cet audit se déroule en deux temps : la **collecte** sur le contrôleur de
domaine (qui a accès à l'AD et au réseau des serveurs), puis la **mise en
forme** sur votre PC (qui a Excel/Python). Cette séparation évite d'installer
Python sur un contrôleur de domaine.

### Étape 1 — Collecte sur le contrôleur de domaine

1. Connectez-vous au **contrôleur de domaine** (ou à une machine avec RSAT et
   accès WinRM aux serveurs), de préférence en **administrateur de domaine**.
2. Copiez `Audit_Securite_Serveurs_DC.ps1` sur la machine, puis ouvrez
   **PowerShell en tant qu'administrateur** dans le dossier du script.
3. Autorisez l'exécution pour la session en cours :

   ```powershell
   Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
   ```

4. Lancez la collecte :

   ```powershell
   # Tout le domaine
   .\Audit_Securite_Serveurs_DC.ps1

   # Cibler une OU précise
   .\Audit_Securite_Serveurs_DC.ps1 -SearchBase "OU=Serveurs,DC=example,DC=local"

   # Ajuster la charge réseau (grands parcs)
   .\Audit_Securite_Serveurs_DC.ps1 -BatchSize 10 -ThrottleLimit 5 -PauseSec 10
   ```

   Paramètres utiles : `-OutputDir` (dossier de sortie, Bureau par défaut),
   `-BatchSize` (serveurs par vague, 20), `-ThrottleLimit` (connexions WinRM
   simultanées, 10), `-PauseSec` (pause entre vagues, 5),
   `-IncludeSystemShares` (inclure `ADMIN$`, `C$`… exclus par défaut).

5. À la fin, trois fichiers sont créés dans le dossier de sortie :
   `Audit_Securite_<date>.csv`, `Audit_Securite_<date>_meta.json` et un
   journal `.log`. Une **reprise automatique** est prévue : si le script est
   interrompu, relancez-le dans les 24 h pour continuer là où il s'était arrêté.

### Étape 2 — Mise en forme sur votre PC

1. Copiez le **CSV** et le **JSON** produits à l'étape 1 sur votre PC (par
   exemple sur le Bureau).
2. Copiez `Formater_Rapport_Excel.ps1` sur votre PC, ouvrez PowerShell dans son
   dossier et lancez :

   ```powershell
   # Détecte automatiquement le CSV le plus récent sur le Bureau
   .\Formater_Rapport_Excel.ps1

   # Ou en pointant explicitement le fichier
   .\Formater_Rapport_Excel.ps1 -CsvFile "C:\Audits\Audit_Securite_<date>.csv"
   ```

3. Si Python 3 ou openpyxl manquent, le script les installe automatiquement
   (via winget puis pip). Le fichier **`.xlsx`** est généré à côté du CSV et
   peut être ouvert directement à la fin.

### Lecture du rapport Excel

Le classeur comporte une feuille **Resume** (synthèse + graphique), une feuille
**Audit Partages** (détail complet), une feuille **par niveau de risque**, et
si besoin une feuille **Serveurs KO** (injoignables) et **Erreurs Audit**.

Les partages exposés sont classés en quatre niveaux :

| Niveau | Signification |
|---|---|
| **CRITIQUE** | Exposé à « Everyone » côté partage **et** NTFS, avec contrôle total |
| **ELEVE** | Exposé à « Everyone » côté partage **et** NTFS |
| **MOYEN** | Exposé à « Everyone » uniquement côté partage |
| **FAIBLE** | Exposé à « Everyone » uniquement côté NTFS |

### Dépannage

- **Serveurs en « WinRM KO »** : activez WinRM (`winrm quickconfig`) et ouvrez
  le port **5985** dans le pare-feu des serveurs concernés.
- **Lectures NTFS limitées** : exécutez le script en administrateur de domaine.
- **Blocage d'exécution** : `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`.
- **Python introuvable** : installez-le depuis <https://www.python.org/downloads/>
  puis relancez `Formater_Rapport_Excel.ps1`.

## Prérequis

- Windows avec le module **RSAT ActiveDirectory**
  (la commande `Get-ADUser` doit être disponible)
- Pour `Export-GPOReport.ps1` : module **RSAT GroupPolicy**
  (la commande `Get-GPO` doit être disponible)
- Module PowerShell **ImportExcel** — installé automatiquement par les scripts
  d'extraction s'il est absent (`Install-Module ImportExcel -Scope CurrentUser`)
- Pour `Formater_Rapport_Excel.ps1` uniquement : **Python 3 + openpyxl**
  (installés automatiquement si absents)
- Droits de **lecture** sur l'annuaire Active Directory (et sur SYSVOL pour
  lister les fichiers de scripts des GPO)

## Utilisation générale

La plupart des scripts d'extraction AD partagent les mêmes options. Par défaut,
le rapport est déposé sur le **Bureau** de l'utilisateur courant.

```powershell
# Tout le domaine
.\Export-ADUsersReport.ps1

# Limiter à une OU précise (commun à la majorité des scripts)
.\Export-ADUsersReport.ps1 -SearchBase "OU=Utilisateurs,DC=example,DC=local"

# Choisir l'emplacement de sortie
.\Export-ADUsersReport.ps1 -OutputPath "C:\Rapports\ad_users.xlsx"
```

Options communes : `-SearchBase` (cibler une OU), `-OutputPath` (fichier de
sortie). Certains scripts ajoutent leurs propres paramètres, décrits dans leur
section ci-dessus et dans l'aide intégrée (`Get-Help .\Script.ps1 -Full`).

> Si l'exécution des scripts est bloquée par la politique de sécurité, lancez :
> `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`

> **Où lancer les scripts ?** De préférence depuis un **contrôleur de domaine**
> (ou une machine avec les outils RSAT), avec des droits de lecture sur l'AD.
> Les scripts de sécurité/inventaire donnent des résultats plus complets exécutés
> au plus près de l'annuaire et du réseau.

## Avertissement

Ces scripts sont fournis à des fins d'administration légitime de votre propre
domaine Active Directory. Utilisez-les uniquement avec les autorisations
appropriées.

## Auteur

Valérian DUFOUR / Claude
