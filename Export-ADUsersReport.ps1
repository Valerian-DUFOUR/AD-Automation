<#
.SYNOPSIS
    Extraction des comptes utilisateurs Active Directory et génération d'un
    rapport Excel (.xlsx) formaté sur le Bureau.

.DESCRIPTION
    Ce script interroge Active Directory via le module ActiveDirectory
    (Get-ADUser) et exporte, pour chaque compte utilisateur, les informations
    suivantes :
        - Nom du compte (SamAccountName)
        - Adresse e-mail
        - Description
        - Nom (Nom de famille / Surname)
        - Prénom (GivenName)
        - Nom complet (DisplayName)
        - Dernière connexion (LastLogonDate)
        - Dernier changement de mot de passe (PasswordLastSet)
        - Date de création du compte (whenCreated)
        - Emplacement dans l'AD (Unité d'Organisation / OU)
        - Compte activé ou désactivé
        - Mot de passe en "n'expire jamais" (Never Expire)

    Le rapport est exporté au format Excel avec un formatage soigné
    (tableau filtrable, en-tête coloré, mise en évidence des comptes
    désactivés et des mots de passe qui n'expirent jamais, ajustement
    automatique des colonnes) grâce au module ImportExcel.

.PARAMETER SearchBase
    (Optionnel) DN de l'OU de départ pour limiter la recherche.
    Exemple : "OU=Utilisateurs,DC=contoso,DC=local"
    Par défaut : tout le domaine.

.PARAMETER OutputPath
    (Optionnel) Chemin complet du fichier .xlsx de sortie.
    Par défaut : le Bureau de l'utilisateur courant.

.EXAMPLE
    .\Export-ADUsersReport.ps1

.EXAMPLE
    .\Export-ADUsersReport.ps1 -SearchBase "OU=Sièges,DC=contoso,DC=local"

.NOTES
    Prérequis :
        - Windows avec le module RSAT ActiveDirectory
          (Get-Command Get-ADUser doit fonctionner)
        - Module ImportExcel (installé automatiquement si absent)
        - Droits de lecture sur l'annuaire Active Directory

    Auteur : généré via Cowork (Claude)
#>

[CmdletBinding()]
param(
    [string]$SearchBase,

    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Utilisateurs_{0}.xlsx" -f (Get-Date -Format 'yyyy-MM-dd_HHmm')))
)

$ErrorActionPreference = 'Stop'

function Write-Step {
    param([string]$Message)
    Write-Host "[+] $Message" -ForegroundColor Cyan
}

# --------------------------------------------------------------------------
# 1. Vérification des prérequis
# --------------------------------------------------------------------------
Write-Step "Vérification du module ActiveDirectory..."
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "Le module 'ActiveDirectory' est introuvable. Installez les outils RSAT (Remote Server Administration Tools) puis relancez le script."
}
Import-Module ActiveDirectory -ErrorAction Stop

Write-Step "Vérification du module ImportExcel..."
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Step "Module ImportExcel absent : installation pour l'utilisateur courant..."
    try {
        Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop
    }
    catch {
        throw "Impossible d'installer le module ImportExcel automatiquement. Exécutez manuellement : Install-Module ImportExcel -Scope CurrentUser"
    }
}
Import-Module ImportExcel -ErrorAction Stop

# --------------------------------------------------------------------------
# 2. Récupération des utilisateurs Active Directory
# --------------------------------------------------------------------------
Write-Step "Interrogation d'Active Directory..."

$properties = @(
    'SamAccountName',
    'mail',
    'Description',
    'Surname',
    'GivenName',
    'DisplayName',
    'LastLogonDate',
    'PasswordLastSet',
    'whenCreated',
    'DistinguishedName',
    'Enabled',
    'PasswordNeverExpires'
)

$getParams = @{
    Filter     = '*'
    Properties = $properties
}
if ($SearchBase) {
    $getParams['SearchBase'] = $SearchBase
    Write-Step "Périmètre limité à : $SearchBase"
}

$users = Get-ADUser @getParams

Write-Step ("{0} compte(s) utilisateur récupéré(s)." -f $users.Count)

# --------------------------------------------------------------------------
# 3. Mise en forme des données
# --------------------------------------------------------------------------
function Get-OUFromDN {
    param([string]$DistinguishedName)
    # Retire le premier composant (CN=...) pour ne garder que l'emplacement (OU/conteneur)
    if ($DistinguishedName -match '^CN=.*?,(.*)$') {
        return $Matches[1]
    }
    return $DistinguishedName
}

Write-Step "Préparation du rapport..."
$report = $users | ForEach-Object {
    [PSCustomObject]@{
        'Nom du compte'            = $_.SamAccountName
        'Email'                    = $_.mail
        'Prénom'                   = $_.GivenName
        'Nom'                      = $_.Surname
        'Nom complet'              = $_.DisplayName
        'Description'              = $_.Description
        'Dernière connexion'       = $_.LastLogonDate
        'Dernier changement MDP'   = $_.PasswordLastSet
        'Date de création'         = $_.whenCreated
        'Emplacement (OU)'         = (Get-OUFromDN $_.DistinguishedName)
        'Statut'                   = if ($_.Enabled) { 'Activé' } else { 'Désactivé' }
        'MDP n''expire jamais'     = if ($_.PasswordNeverExpires) { 'Oui' } else { 'Non' }
    }
} | Sort-Object 'Nom du compte'

# --------------------------------------------------------------------------
# 4. Export Excel formaté
# --------------------------------------------------------------------------
Write-Step "Génération du fichier Excel : $OutputPath"

# Suppression d'un éventuel fichier existant portant le même nom
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

$excelParams = @{
    Path          = $OutputPath
    WorksheetName = 'Utilisateurs AD'
    AutoSize      = $true
    AutoFilter    = $true
    FreezeTopRow  = $true
    BoldTopRow    = $true
    TableName     = 'UtilisateursAD'
    TableStyle    = 'Medium2'
    Title         = ("Rapport des comptes Active Directory - {0}" -f (Get-Date -Format 'dd/MM/yyyy HH:mm'))
    TitleBold     = $true
    TitleSize     = 14
    PassThru      = $true
}

$excel = $report | Export-Excel @excelParams

# Mise en forme conditionnelle
$ws = $excel.Workbook.Worksheets['Utilisateurs AD']

# Colonnes contenant des dates -> format lisible
$dateColumns = @('Dernière connexion', 'Dernier changement MDP', 'Date de création')
$headerRow   = $ws.Dimension.Start.Row  # ligne d'en-tête (sous le titre)
foreach ($colName in $dateColumns) {
    for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
        if ($ws.Cells[$headerRow, $c].Value -eq $colName) {
            $colLetter = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c)
            $range = "{0}{1}:{0}{2}" -f $colLetter, ($headerRow + 1), $ws.Dimension.End.Row
            $ws.Cells[$range].Style.Numberformat.Format = 'yyyy-mm-dd hh:mm'
        }
    }
}

# Surligner en rouge clair les comptes désactivés + les MDP qui n'expirent jamais
$lastRow       = $ws.Dimension.End.Row
$firstDataRow  = $headerRow + 1
$lastColLetter = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column)
$dataRange     = "A{0}:{1}{2}" -f $firstDataRow, $lastColLetter, $lastRow

# Les expressions sont relatives à la 1ère cellule de la plage (ligne $firstDataRow)
Add-ConditionalFormatting -Worksheet $ws -Range $dataRange -RuleType Expression `
    -ConditionValue ('=$K{0}="Désactivé"' -f $firstDataRow) -BackgroundColor ([System.Drawing.Color]::MistyRose)

Add-ConditionalFormatting -Worksheet $ws -Range $dataRange -RuleType Expression `
    -ConditionValue ('=$L{0}="Oui"' -f $firstDataRow) -ForegroundColor ([System.Drawing.Color]::DarkRed) -Bold

Close-ExcelPackage $excel

Write-Host ""
Write-Host "[OK] Rapport généré avec succès : $OutputPath" -ForegroundColor Green
Write-Host ("     {0} compte(s) exporté(s)." -f $report.Count) -ForegroundColor Green
