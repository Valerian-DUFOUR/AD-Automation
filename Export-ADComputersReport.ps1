<#
.SYNOPSIS
    Extraction des ordinateurs Active Directory et génération d'un rapport
    Excel (.xlsx) formaté sur le Bureau, avec une page de synthèse (KPI).

.DESCRIPTION
    Ce script interroge Active Directory via le module ActiveDirectory
    (Get-ADComputer) et exporte, pour chaque ordinateur, les informations
    suivantes :
        - Nom du PC
        - Description
        - Système d'exploitation (OS)
        - Version de l'OS
        - Date de création du compte machine
        - Date de dernière connexion
        - Compte activé ou désactivé dans l'AD
        - Nombre de jours d'inactivité
        - Emplacement dans l'AD (Unité d'Organisation / OU)
        - Obsolescence : indique si le poste est obsolète selon sa version
          de Windows (fin de support)

    Le classeur Excel généré contient DEUX feuilles :
        1. « Synthèse » : indicateurs clés (KPI) de l'extraction
           (total, activés/désactivés, obsolètes, inactifs, répartition
           par OS) avec un graphique.
        2. « Ordinateurs » : le détail de tous les postes, sous forme de
           tableau filtrable et mis en forme (postes obsolètes, désactivés
           et inactifs mis en évidence).

.PARAMETER SearchBase
    (Optionnel) DN de l'OU de départ pour limiter la recherche.
    Exemple : "OU=Ordinateurs,DC=contoso,DC=local"
    Par défaut : tout le domaine.

.PARAMETER InactiveDays
    (Optionnel) Seuil, en jours, à partir duquel un poste est considéré
    comme inactif. Par défaut : 90 jours.

.PARAMETER OutputPath
    (Optionnel) Chemin complet du fichier .xlsx de sortie.
    Par défaut : le Bureau de l'utilisateur courant.

.EXAMPLE
    .\Export-ADComputersReport.ps1

.EXAMPLE
    .\Export-ADComputersReport.ps1 -SearchBase "OU=Postes,DC=contoso,DC=local" -InactiveDays 60

.NOTES
    Prérequis :
        - Windows avec le module RSAT ActiveDirectory
          (Get-Command Get-ADComputer doit fonctionner)
        - Module ImportExcel (installé automatiquement si absent)
        - Droits de lecture sur l'annuaire Active Directory

    Auteur : Valérian DUFOUR / Claude
#>

[CmdletBinding()]
param(
    [string]$SearchBase,

    [int]$InactiveDays = 90,

    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Ordinateurs_{0}.xlsx" -f (Get-Date -Format 'yyyy-MM-dd_HHmm')))
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
# 2. Fonctions utilitaires
# --------------------------------------------------------------------------
function Get-OUFromDN {
    param([string]$DistinguishedName)
    if ($DistinguishedName -match '^CN=.*?,(.*)$') { return $Matches[1] }
    return $DistinguishedName
}

function Get-ObsolescenceStatus {
    <#
        Détermine si un poste est obsolète selon son OS (fin de support Microsoft).
        Contexte : Windows 10 est en fin de support depuis le 14/10/2025.
        Adaptez cette logique selon votre politique interne.
    #>
    param([string]$OperatingSystem)

    if ([string]::IsNullOrWhiteSpace($OperatingSystem)) {
        return [PSCustomObject]@{ Obsolete = 'Inconnu'; Motif = 'OS non renseigné' }
    }

    $os = $OperatingSystem

    # --- Systèmes clients ---
    if ($os -match 'Windows 11')      { return [PSCustomObject]@{ Obsolete = 'Non'; Motif = 'Support en cours' } }
    if ($os -match 'Windows 10')      { return [PSCustomObject]@{ Obsolete = 'Oui'; Motif = 'Fin de support (14/10/2025)' } }
    if ($os -match 'Windows 8')       { return [PSCustomObject]@{ Obsolete = 'Oui'; Motif = 'Fin de support (Windows 8/8.1)' } }
    if ($os -match 'Windows 7')       { return [PSCustomObject]@{ Obsolete = 'Oui'; Motif = 'Fin de support (14/01/2020)' } }
    if ($os -match 'Windows Vista')   { return [PSCustomObject]@{ Obsolete = 'Oui'; Motif = 'Fin de support' } }
    if ($os -match 'Windows XP')      { return [PSCustomObject]@{ Obsolete = 'Oui'; Motif = 'Fin de support' } }

    # --- Systèmes serveurs ---
    if ($os -match 'Server 2025')     { return [PSCustomObject]@{ Obsolete = 'Non'; Motif = 'Support en cours' } }
    if ($os -match 'Server 2022')     { return [PSCustomObject]@{ Obsolete = 'Non'; Motif = 'Support en cours' } }
    if ($os -match 'Server 2019')     { return [PSCustomObject]@{ Obsolete = 'Non'; Motif = 'Support en cours' } }
    if ($os -match 'Server 2016')     { return [PSCustomObject]@{ Obsolete = 'Non'; Motif = 'Support étendu (jusqu''au 12/01/2027)' } }
    if ($os -match 'Server 2012')     { return [PSCustomObject]@{ Obsolete = 'Oui'; Motif = 'Fin de support (10/10/2023)' } }
    if ($os -match 'Server 2008')     { return [PSCustomObject]@{ Obsolete = 'Oui'; Motif = 'Fin de support (14/01/2020)' } }
    if ($os -match 'Server 2003')     { return [PSCustomObject]@{ Obsolete = 'Oui'; Motif = 'Fin de support' } }

    return [PSCustomObject]@{ Obsolete = 'À vérifier'; Motif = 'OS non répertorié' }
}

# --------------------------------------------------------------------------
# 3. Récupération des ordinateurs Active Directory
# --------------------------------------------------------------------------
Write-Step "Interrogation d'Active Directory..."

$properties = @(
    'Name',
    'Description',
    'OperatingSystem',
    'OperatingSystemVersion',
    'whenCreated',
    'LastLogonDate',
    'Enabled',
    'DistinguishedName'
)

$getParams = @{
    Filter     = '*'
    Properties = $properties
}
if ($SearchBase) {
    $getParams['SearchBase'] = $SearchBase
    Write-Step "Périmètre limité à : $SearchBase"
}

$computers = Get-ADComputer @getParams
Write-Step ("{0} ordinateur(s) récupéré(s)." -f $computers.Count)

# --------------------------------------------------------------------------
# 4. Mise en forme des données
# --------------------------------------------------------------------------
Write-Step "Préparation du rapport..."
$now = Get-Date

$report = $computers | ForEach-Object {
    $daysInactive = if ($_.LastLogonDate) {
        [int]($now - $_.LastLogonDate).TotalDays
    } else {
        $null
    }

    $obs = Get-ObsolescenceStatus -OperatingSystem $_.OperatingSystem

    [PSCustomObject]@{
        'Nom du PC'              = $_.Name
        'Description'            = $_.Description
        'OS'                     = $_.OperatingSystem
        'Version OS'             = $_.OperatingSystemVersion
        'Date de création'       = $_.whenCreated
        'Dernière connexion'     = $_.LastLogonDate
        'Statut'                 = if ($_.Enabled) { 'Activé' } else { 'Désactivé' }
        'Jours inactif'          = $daysInactive
        'Emplacement (OU)'       = (Get-OUFromDN $_.DistinguishedName)
        'Obsolète'               = $obs.Obsolete
        'Motif obsolescence'     = $obs.Motif
    }
} | Sort-Object 'Nom du PC'

# --------------------------------------------------------------------------
# 5. Calcul des indicateurs (KPI)
# --------------------------------------------------------------------------
$total       = @($report).Count
$actifs      = @($report | Where-Object { $_.Statut -eq 'Activé' }).Count
$desactives  = @($report | Where-Object { $_.Statut -eq 'Désactivé' }).Count
$obsoletes   = @($report | Where-Object { $_.'Obsolète' -eq 'Oui' }).Count
$inactifs    = @($report | Where-Object { $_.'Jours inactif' -ne $null -and $_.'Jours inactif' -ge $InactiveDays }).Count

$pct = { param($n) if ($total -gt 0) { [math]::Round(($n / $total) * 100, 1) } else { 0 } }

# Répartition par OS
$osDistribution = $report |
    Group-Object 'OS' |
    Sort-Object Count -Descending |
    ForEach-Object {
        [PSCustomObject]@{
            'Système d''exploitation' = if ($_.Name) { $_.Name } else { '(non renseigné)' }
            'Nombre'                  = $_.Count
        }
    }

# --------------------------------------------------------------------------
# 6. Génération du fichier Excel
# --------------------------------------------------------------------------
Write-Step "Génération du fichier Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

# --- Feuille de détail « Ordinateurs » ---
$excel = $report | Export-Excel -Path $OutputPath -WorksheetName 'Ordinateurs' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName 'Ordinateurs' -TableStyle 'Medium2' `
    -Title ("Détail des ordinateurs Active Directory - {0}" -f $now.ToString('dd/MM/yyyy HH:mm')) `
    -TitleBold -TitleSize 14 -PassThru

$wsData    = $excel.Workbook.Worksheets['Ordinateurs']
$headerRow = $wsData.Dimension.Start.Row
$firstData = $headerRow + 1
$lastRow   = $wsData.Dimension.End.Row
$lastCol   = $wsData.Dimension.End.Column
$lastColL  = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($lastCol)

# Format des colonnes de dates
foreach ($colName in @('Date de création', 'Dernière connexion')) {
    for ($c = 1; $c -le $lastCol; $c++) {
        if ($wsData.Cells[$headerRow, $c].Value -eq $colName) {
            $l = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c)
            $wsData.Cells[("{0}{1}:{0}{2}" -f $l, $firstData, $lastRow)].Style.Numberformat.Format = 'yyyy-mm-dd hh:mm'
        }
    }
}

$dataRange = "A{0}:{1}{2}" -f $firstData, $lastColL, $lastRow

# Colonne J = Obsolète, G = Statut, H = Jours inactif
Add-ConditionalFormatting -Worksheet $wsData -Range $dataRange -RuleType Expression `
    -ConditionValue ('=$J{0}="Oui"' -f $firstData) -BackgroundColor ([System.Drawing.Color]::MistyRose)

Add-ConditionalFormatting -Worksheet $wsData -Range $dataRange -RuleType Expression `
    -ConditionValue ('=$G{0}="Désactivé"' -f $firstData) -ForegroundColor ([System.Drawing.Color]::Gray) -Italic

# Mise en évidence de la colonne « Jours inactif » au-delà du seuil
for ($c = 1; $c -le $lastCol; $c++) {
    if ($wsData.Cells[$headerRow, $c].Value -eq 'Jours inactif') {
        $l = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c)
        $inactRange = "{0}{1}:{0}{2}" -f $l, $firstData, $lastRow
        Add-ConditionalFormatting -Worksheet $wsData -Range $inactRange -RuleType GreaterThanOrEqual `
            -ConditionValue $InactiveDays -ForegroundColor ([System.Drawing.Color]::DarkRed) -Bold
    }
}

# --- Feuille de synthèse « Synthèse » (placée en 1ère position) ---
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse' -MoveToStart

# Titre
$wsK.Cells['B2'].Value = 'Rapport des ordinateurs Active Directory'
$wsK.Cells['B2'].Style.Font.Size = 16
$wsK.Cells['B2'].Style.Font.Bold = $true
$wsK.Cells['B2:E2'].Merge = $true

$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f $now.ToString('dd/MM/yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$wsK.Cells['B5'].Value = ("Seuil d'inactivité : {0} jours" -f $InactiveDays)

# Bloc KPI
$kpis = @(
    @('Total des postes',          $total,      $null),
    @('Postes activés',            $actifs,     (& $pct $actifs)),
    @('Postes désactivés',         $desactives, (& $pct $desactives)),
    @('Postes obsolètes',          $obsoletes,  (& $pct $obsoletes)),
    @("Postes inactifs (>= $InactiveDays j)", $inactifs, (& $pct $inactifs))
)

$row = 7
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
    # Alternance de couleur
    if ($row % 2 -eq 0) {
        $wsK.Cells["B$row:D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217,226,243))
    }
}
$kpiLastRow = $row

# Mise en évidence du nombre de postes obsolètes en rouge (4e ligne de KPI)
$obsoleteKpiRow = 7 + 4
$wsK.Cells["C$obsoleteKpiRow"].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
$wsK.Cells["C$obsoleteKpiRow"].Style.Font.Bold = $true

# Bloc répartition par OS
$osHeaderRow = $kpiLastRow + 3
$wsK.Cells["B$osHeaderRow"].Value = 'Système d''exploitation'
$wsK.Cells["C$osHeaderRow"].Value = 'Nombre'
$wsK.Cells["B$osHeaderRow:C$osHeaderRow"].Style.Font.Bold = $true
$wsK.Cells["B$osHeaderRow:C$osHeaderRow"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$osHeaderRow:C$osHeaderRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$osHeaderRow:C$osHeaderRow"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

$r = $osHeaderRow
foreach ($os in $osDistribution) {
    $r++
    $wsK.Cells["B$r"].Value = $os.'Système d''exploitation'
    $wsK.Cells["C$r"].Value = $os.'Nombre'
}
$osLastRow = $r

# Graphique de répartition par OS
try {
    $chart = $wsK.Drawings.AddChart('osChart', [OfficeOpenXml.Drawing.Chart.eChartType]::BarClustered)
    $chart.Title.Text = 'Répartition par système d''exploitation'
    $chart.SetPosition(6, 0, 5, 0)   # ligne, offset, colonne, offset
    $chart.SetSize(520, 300)
    $serie = $chart.Series.Add(
        $wsK.Cells["C$($osHeaderRow + 1):C$osLastRow"],
        $wsK.Cells["B$($osHeaderRow + 1):B$osLastRow"]
    )
    $serie.Header = 'Nombre de postes'
    $chart.DataLabel.ShowValue = $true
}
catch {
    Write-Warning "Graphique non généré (API EPPlus indisponible) : $($_.Exception.Message)"
}

# Ajustement des largeurs de colonnes de la feuille de synthèse
$wsK.Column(2).Width = 34
$wsK.Column(3).Width = 14
$wsK.Column(4).Width = 10

Close-ExcelPackage $excel

Write-Host ""
Write-Host "[OK] Rapport généré avec succès : $OutputPath" -ForegroundColor Green
Write-Host ("     {0} poste(s) | {1} obsolète(s) | {2} inactif(s) | {3} désactivé(s)." -f `
    $total, $obsoletes, $inactifs, $desactives) -ForegroundColor Green
