# AD Automation — Scripts d'extraction Active Directory

Collection de scripts PowerShell pour automatiser l'extraction d'informations
Active Directory et générer des rapports Excel formatés.

## Scripts disponibles

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

Mise en forme du rapport :

- Tableau filtrable avec en-tête en gras et première ligne figée
- Colonnes ajustées automatiquement
- Dates formatées `AAAA-MM-JJ HH:MM`
- Comptes **désactivés** surlignés en rouge clair
- Mots de passe en **« n'expire jamais »** signalés en rouge gras

## Prérequis

- Windows avec le module **RSAT ActiveDirectory**
  (la commande `Get-ADUser` doit être disponible)
- Module PowerShell **ImportExcel** — installé automatiquement par le script
  s'il est absent (`Install-Module ImportExcel -Scope CurrentUser`)
- Droits de **lecture** sur l'annuaire Active Directory

## Utilisation

Extraction de tout le domaine, rapport déposé sur le Bureau :

```powershell
.\Export-ADUsersReport.ps1
```

Limiter la recherche à une OU précise :

```powershell
.\Export-ADUsersReport.ps1 -SearchBase "OU=Utilisateurs,DC=contoso,DC=local"
```

Choisir un emplacement de sortie personnalisé :

```powershell
.\Export-ADUsersReport.ps1 -OutputPath "C:\Rapports\ad_users.xlsx"
```

> Si l'exécution des scripts est bloquée par la politique de sécurité, lancez :
> `Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass`

## Avertissement

Ces scripts sont fournis à des fins d'administration légitime de votre propre
domaine Active Directory. Utilisez-les uniquement avec les autorisations
appropriées.
