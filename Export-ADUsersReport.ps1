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

    Le classeur Excel généré contient DEUX feuilles :
        1. « Synthèse » : indicateurs clés (KPI) de l'extraction
           (total, activés/désactivés, MDP n'expire jamais, comptes sans
           e-mail), répartition par OU et un graphique.
        2. « Utilisateurs AD » : le détail de tous les comptes, sous forme
           de tableau filtrable et mis en forme (comptes désactivés et mots
           de passe qui n'expirent jamais mis en évidence).

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

    Auteur : Valérian DUFOUR / Claude
#>

[CmdletBinding()]
param(
    [string]$SearchBase,

    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Utilisateurs_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
)

$ErrorActionPreference = 'Stop'
$Author = 'Valérian DUFOUR / Claude'

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
$now = Get-Date
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
    Title         = ("Rapport des comptes Active Directory - {0}" -f $now.ToString('dd-MM-yyyy HH:mm'))
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
            $ws.Cells[$range].Style.Numberformat.Format = 'dd-mm-yyyy hh:mm'
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

# --------------------------------------------------------------------------
# 5. Calcul des indicateurs (KPI)
# --------------------------------------------------------------------------
$total       = @($report).Count
$actifs      = @($report | Where-Object { $_.Statut -eq 'Activé' }).Count
$desactives  = @($report | Where-Object { $_.Statut -eq 'Désactivé' }).Count
$neverExp    = @($report | Where-Object { $_.'MDP n''expire jamais' -eq 'Oui' }).Count
$sansEmail   = @($report | Where-Object { [string]::IsNullOrWhiteSpace($_.Email) }).Count

$pct = { param($n) if ($total -gt 0) { [math]::Round(($n / $total) * 100, 1) } else { 0 } }

# Répartition par emplacement (OU)
$ouDistribution = $report |
    Group-Object 'Emplacement (OU)' |
    Sort-Object Count -Descending |
    ForEach-Object {
        [PSCustomObject]@{
            'Emplacement (OU)' = if ($_.Name) { $_.Name } else { '(non renseigné)' }
            'Nombre'           = $_.Count
        }
    }

# --------------------------------------------------------------------------
# 6. Feuille de synthèse « Synthèse » (placée en 1ère position)
# --------------------------------------------------------------------------
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse' -MoveToStart

# Titre
$wsK.Cells['B2'].Value = 'Rapport des comptes utilisateurs Active Directory'
$wsK.Cells['B2'].Style.Font.Size = 16
$wsK.Cells['B2'].Style.Font.Bold = $true
$wsK.Cells['B2:E2'].Merge = $true

$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f $now.ToString('dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)

# Bloc KPI
$kpis = @(
    @('Total des comptes',            $total,      $null),
    @('Comptes activés',              $actifs,     (& $pct $actifs)),
    @('Comptes désactivés',           $desactives, (& $pct $desactives)),
    @("MDP n'expire jamais",          $neverExp,   (& $pct $neverExp)),
    @('Comptes sans e-mail',          $sansEmail,  (& $pct $sansEmail))
)

$row = 6
$wsK.Cells["B$row"].Value = 'Indicateur'
$wsK.Cells["C$row"].Value = 'Valeur'
$wsK.Cells["D$row"].Value = '%'
$wsK.Cells["B$row:D$row"].Style.Font.Bold = $true
$wsK.Cells["B$row:D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$row:D$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

foreach ($k in $kpis) {
    $row++
    $wsK.Cells["B$row"].Value = $k[0]
    $wsK.Cells["C$row"].Value = $k[1]
    if ($null -ne $k[2]) { $wsK.Cells["D$row"].Value = ("{0} %" -f $k[2]) }
    if ($row % 2 -eq 0) {
        $wsK.Cells["B$row:D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217,226,243))
    }
}
$kpiLastRow = $row

# Mise en évidence du nombre de comptes désactivés en rouge (3e ligne de KPI)
$desactivesKpiRow = 6 + 3
$wsK.Cells["C$desactivesKpiRow"].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
$wsK.Cells["C$desactivesKpiRow"].Style.Font.Bold = $true

# Bloc répartition par OU
$ouHeaderRow = $kpiLastRow + 3
$wsK.Cells["B$ouHeaderRow"].Value = 'Emplacement (OU)'
$wsK.Cells["C$ouHeaderRow"].Value = 'Nombre'
$wsK.Cells["B$ouHeaderRow:C$ouHeaderRow"].Style.Font.Bold = $true
$wsK.Cells["B$ouHeaderRow:C$ouHeaderRow"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$ouHeaderRow:C$ouHeaderRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$ouHeaderRow:C$ouHeaderRow"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

$r = $ouHeaderRow
foreach ($ou in $ouDistribution) {
    $r++
    $wsK.Cells["B$r"].Value = $ou.'Emplacement (OU)'
    $wsK.Cells["C$r"].Value = $ou.'Nombre'
}
$ouLastRow = $r

# Graphique de répartition par statut (activés / désactivés / never expire)
try {
    $wsK.Cells['F6'].Value  = 'Répartition'
    $wsK.Cells['F7'].Value  = 'Activés';            $wsK.Cells['G7'].Value = $actifs
    $wsK.Cells['F8'].Value  = 'Désactivés';         $wsK.Cells['G8'].Value = $desactives
    $wsK.Cells['F9'].Value  = "MDP n'expire jamais"; $wsK.Cells['G9'].Value = $neverExp
    $wsK.Cells['F6:G6'].Style.Font.Bold = $true

    $chart = $wsK.Drawings.AddChart('userChart', [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
    $chart.Title.Text = 'Répartition des comptes'
    $chart.SetPosition(5, 0, 8, 0)   # ligne, offset, colonne, offset
    $chart.SetSize(420, 300)
    $serie = $chart.Series.Add($wsK.Cells['G7:G9'], $wsK.Cells['F7:F9'])
    $chart.DataLabel.ShowPercent = $true
}
catch {
    Write-Warning "Graphique non généré (API EPPlus indisponible) : $($_.Exception.Message)"
}

# Ajustement des largeurs de colonnes de la feuille de synthèse
$wsK.Column(2).Width = 42
$wsK.Column(3).Width = 14
$wsK.Column(4).Width = 10

Close-ExcelPackage $excel

Write-Host ""
Write-Host "[OK] Rapport généré avec succès : $OutputPath" -ForegroundColor Green
Write-Host ("     {0} compte(s) exporté(s)." -f $report.Count) -ForegroundColor Green
