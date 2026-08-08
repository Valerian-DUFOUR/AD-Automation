<#
.SYNOPSIS
    Export des ordinateurs Active Directory NE disposant PAS d'une clé de
    récupération BitLocker sauvegardée dans l'AD, avec page KPI, en Excel
    (.xlsx) sur le Bureau.

.DESCRIPTION
    Reprend les mêmes informations que le rapport des ordinateurs
    (Export-ADComputersReport.ps1) mais ne conserve que les postes SANS clé
    BitLocker escrowée dans l'annuaire.

    Méthode de détection : un poste dont la clé BitLocker est sauvegardée dans
    l'AD possède un objet enfant de classe « msFVE-RecoveryInformation ». Ce
    script recense tous ces objets, en déduit les postes couverts, et liste
    ceux qui ne le sont pas.

    IMPORTANT : ceci vérifie la présence d'une clé de récupération DANS L'AD,
    et non le statut de chiffrement réel du disque. Un poste peut être chiffré
    sans que sa clé soit sauvegardée dans l'AD (et inversement, rare). Pour le
    statut de chiffrement réel, il faut interroger la machine
    (Get-BitLockerVolume / manage-bde), ce qui sort du périmètre de ce script.

    Le classeur contient :
        - « Synthèse (KPI) » : total de postes, postes avec / sans clé,
          taux de couverture, et graphique.
        - « PC sans BitLocker » : le détail (mêmes colonnes que l'export PC).

    Toutes les dates sont au format JJ-MM-AAAA.

.PARAMETER SearchBase
    (Optionnel) DN de l'OU de départ. Par défaut : tout le domaine.

.PARAMETER OutputPath
    (Optionnel) Chemin complet du .xlsx. Par défaut : le Bureau.

.EXAMPLE
    .\Export-ADComputersNoBitLocker.ps1

.NOTES
    Prérequis :
        - RSAT ActiveDirectory + module ImportExcel (auto-installé)
        - Droits de lecture sur les objets msFVE-RecoveryInformation
          (généralement réservés aux administrateurs)

    Auteur : Valérian DUFOUR / Claude
#>

[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_PC_Sans_BitLocker_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
)

$ErrorActionPreference = 'Stop'
$Author = 'Valérian DUFOUR / Claude'

function Write-Step { param([string]$Message) Write-Host "[+] $Message" -ForegroundColor Cyan }

# --------------------------------------------------------------------------
# 1. Prérequis
# --------------------------------------------------------------------------
Write-Step "Vérification du module ActiveDirectory..."
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "Le module 'ActiveDirectory' est introuvable. Installez les outils RSAT puis relancez."
}
Import-Module ActiveDirectory -ErrorAction Stop

Write-Step "Vérification du module ImportExcel..."
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Step "Installation du module ImportExcel (utilisateur courant)..."
    try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop }
    catch { throw "Impossible d'installer ImportExcel. Exécutez : Install-Module ImportExcel -Scope CurrentUser" }
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
    param([string]$OperatingSystem)
    if ([string]::IsNullOrWhiteSpace($OperatingSystem)) { return [PSCustomObject]@{ Obsolete='Inconnu'; Motif='OS non renseigné' } }
    $os = $OperatingSystem
    if ($os -match 'Windows 11')    { return [PSCustomObject]@{ Obsolete='Non'; Motif='Support en cours' } }
    if ($os -match 'Windows 10')    { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support (14/10/2025)' } }
    if ($os -match 'Windows 8')     { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support (Windows 8/8.1)' } }
    if ($os -match 'Windows 7')     { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support (14/01/2020)' } }
    if ($os -match 'Windows Vista') { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support' } }
    if ($os -match 'Windows XP')    { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support' } }
    if ($os -match 'Server 2025')   { return [PSCustomObject]@{ Obsolete='Non'; Motif='Support en cours' } }
    if ($os -match 'Server 2022')   { return [PSCustomObject]@{ Obsolete='Non'; Motif='Support en cours' } }
    if ($os -match 'Server 2019')   { return [PSCustomObject]@{ Obsolete='Non'; Motif='Support en cours' } }
    if ($os -match 'Server 2016')   { return [PSCustomObject]@{ Obsolete='Non'; Motif='Support étendu (jusqu''au 12/01/2027)' } }
    if ($os -match 'Server 2012')   { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support (10/10/2023)' } }
    if ($os -match 'Server 2008')   { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support (14/01/2020)' } }
    if ($os -match 'Server 2003')   { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support' } }
    return [PSCustomObject]@{ Obsolete='À vérifier'; Motif='OS non répertorié' }
}

# --------------------------------------------------------------------------
# 3. Collecte
# --------------------------------------------------------------------------
$now = Get-Date
$baseParam = @{}
if ($SearchBase) { $baseParam['SearchBase'] = $SearchBase; Write-Step "Périmètre : $SearchBase" }

Write-Step "Chargement des ordinateurs..."
$computers = Get-ADComputer -Filter * -Properties Name, Description, OperatingSystem, OperatingSystemVersion, `
    whenCreated, LastLogonDate, DistinguishedName, Enabled @baseParam
Write-Step ("{0} ordinateur(s)." -f $computers.Count)

# -- Recensement des postes disposant d'une clé BitLocker dans l'AD --
Write-Step "Recherche des clés de récupération BitLocker (msFVE-RecoveryInformation)..."
$withKeyDN = [System.Collections.Generic.HashSet[string]]::new()
try {
    $recovery = Get-ADObject -LDAPFilter '(objectClass=msFVE-RecoveryInformation)' -Properties distinguishedName @baseParam
    foreach ($r in @($recovery)) {
        # Le parent (le compte ordinateur) = DN sans le premier RDN
        $parent = $r.DistinguishedName -replace '^CN=[^,]+,', ''
        [void]$withKeyDN.Add($parent)
    }
    Write-Step ("{0} clé(s) BitLocker trouvée(s) dans l'AD ({1} poste(s) couvert(s))." -f @($recovery).Count, $withKeyDN.Count)
} catch {
    Write-Warning "Lecture des objets msFVE-RecoveryInformation impossible : $($_.Exception.Message)"
    Write-Warning "Vérifiez que vous disposez des droits nécessaires (administrateur)."
}

# --------------------------------------------------------------------------
# 4. Préparation - postes SANS clé BitLocker
# --------------------------------------------------------------------------
$total = @($computers).Count
$noBitlocker = foreach ($c in $computers) {
    if ($withKeyDN.Contains($c.DistinguishedName)) { continue }
    $days = if ($c.LastLogonDate) { [int]($now - $c.LastLogonDate).TotalDays } else { $null }
    $obs = Get-ObsolescenceStatus -OperatingSystem $c.OperatingSystem
    [PSCustomObject]@{
        'Nom du PC'          = $c.Name
        'Description'        = $c.Description
        'OS'                 = $c.OperatingSystem
        'Version OS'         = $c.OperatingSystemVersion
        'Date de création'   = $c.whenCreated
        'Dernière connexion' = $c.LastLogonDate
        'Statut'             = if ($c.Enabled) { 'Activé' } else { 'Désactivé' }
        'Jours inactif'      = $days
        'Emplacement (OU)'   = (Get-OUFromDN $c.DistinguishedName)
        'Obsolète'           = $obs.Obsolete
        'Motif obsolescence' = $obs.Motif
        'Clé BitLocker (AD)' = 'Absente'
    }
}
$noBitlocker = @($noBitlocker) | Sort-Object 'Nom du PC'

$withCount = @($computers | Where-Object { $withKeyDN.Contains($_.DistinguishedName) }).Count
$withoutCount = $noBitlocker.Count
$coverage = if ($total -gt 0) { [math]::Round(($withCount / $total) * 100, 1) } else { 0 }
Write-Step ("{0} poste(s) sans clé BitLocker sur {1} ({2}% couverts)." -f $withoutCount, $total, $coverage)

# Répartition des postes sans clé par OS
$osDist = $noBitlocker | Group-Object 'OS' | Sort-Object Count -Descending |
    ForEach-Object { [PSCustomObject]@{ 'OS' = if ($_.Name) { $_.Name } else { '(non renseigné)' }; 'Nombre' = $_.Count } }

# --------------------------------------------------------------------------
# 5. Export Excel
# --------------------------------------------------------------------------
Write-Step "Génération du fichier Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

$dataN = if ($withoutCount -gt 0) { $noBitlocker } else { ,([PSCustomObject]@{ 'Information' = 'Tous les postes disposent d''une clé BitLocker dans l''AD' }) }
$excel = $dataN | Export-Excel -Path $OutputPath -WorksheetName 'PC sans BitLocker' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'PCSansBitLocker' -TableStyle 'Medium2' `
    -Title ("Postes sans clé BitLocker dans l'AD - {0}" -f $now.ToString('dd-MM-yyyy HH:mm')) -TitleBold -TitleSize 14 -PassThru

if ($withoutCount -gt 0) {
    $wsN = $excel.Workbook.Worksheets['PC sans BitLocker']
    $hr = $wsN.Dimension.Start.Row
    for ($c = 1; $c -le $wsN.Dimension.End.Column; $c++) {
        if (@('Date de création','Dernière connexion') -contains $wsN.Cells[$hr, $c].Value) {
            $l = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c)
            $wsN.Cells[("{0}{1}:{0}{2}" -f $l, ($hr + 1), $wsN.Dimension.End.Row)].Style.Numberformat.Format = 'dd-mm-yyyy hh:mm'
        }
    }
    # Colonne J = Obsolète
    $rg = "A{0}:{1}{2}" -f ($hr + 1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($wsN.Dimension.End.Column), $wsN.Dimension.End.Row
    Add-ConditionalFormatting -Worksheet $wsN -Range $rg -RuleType Expression `
        -ConditionValue ('=$J{0}="Oui"' -f ($hr + 1)) -BackgroundColor ([System.Drawing.Color]::MistyRose)
}

# --- Feuille Synthèse (KPI) ---
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Couverture BitLocker (clés dans l''AD)'
$wsK.Cells['B2'].Style.Font.Size = 16
$wsK.Cells['B2'].Style.Font.Bold = $true
$wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f $now.ToString('dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$wsK.Cells['B5'].Value = "Détection : présence d'une clé de récupération (msFVE-RecoveryInformation) dans l'AD."

$kpis = @(
    @('Total des postes',            $total,        $null),
    @('Postes avec clé BitLocker',   $withCount,    (if ($total) { [math]::Round(($withCount/$total)*100,1) } else { 0 })),
    @('Postes SANS clé BitLocker',   $withoutCount, (if ($total) { [math]::Round(($withoutCount/$total)*100,1) } else { 0 })),
    @('Taux de couverture (%)',      $coverage,     $null)
)

$row = 7
$wsK.Cells["B$row"].Value = 'Indicateur'; $wsK.Cells["C$row"].Value = 'Valeur'; $wsK.Cells["D$row"].Value = '%'
$wsK.Cells["B$row:D$row"].Style.Font.Bold = $true
$wsK.Cells["B$row:D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$row:D$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) {
    $row++
    $wsK.Cells["B$row"].Value = $k[0]; $wsK.Cells["C$row"].Value = $k[1]
    if ($null -ne $k[2]) { $wsK.Cells["D$row"].Value = ("{0} %" -f $k[2]) }
    if ($row % 2 -eq 0) {
        $wsK.Cells["B$row:D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217,226,243))
    }
}
# Postes sans clé en rouge (3e ligne KPI)
$wsK.Cells["C10"].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
$wsK.Cells["C10"].Style.Font.Bold = $true

# Mini-table avec / sans pour le graphique
$gRow = $row + 3
$wsK.Cells["B$gRow"].Value = 'État'; $wsK.Cells["C$gRow"].Value = 'Nombre'
$wsK.Cells["B$gRow:C$gRow"].Style.Font.Bold = $true
$wsK.Cells["B$($gRow+1)"].Value = 'Avec clé';  $wsK.Cells["C$($gRow+1)"].Value = $withCount
$wsK.Cells["B$($gRow+2)"].Value = 'Sans clé';  $wsK.Cells["C$($gRow+2)"].Value = $withoutCount

try {
    $chart = $wsK.Drawings.AddChart('blChart', [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
    $chart.Title.Text = 'Couverture BitLocker'
    $chart.SetPosition(6, 0, 5, 0); $chart.SetSize(420, 300)
    $null = $chart.Series.Add($wsK.Cells["C$($gRow+1):C$($gRow+2)"], $wsK.Cells["B$($gRow+1):B$($gRow+2)"])
    $chart.DataLabel.ShowPercent = $true
} catch { Write-Warning "Graphique non généré : $($_.Exception.Message)" }

$wsK.Column(2).Width = 30; $wsK.Column(3).Width = 16; $wsK.Column(4).Width = 10

Close-ExcelPackage $excel

Write-Host ""
Write-Host "[OK] Rapport généré : $OutputPath" -ForegroundColor Green
Write-Host ("     {0} poste(s) sans clé BitLocker sur {1} (couverture {2}%)." -f $withoutCount, $total, $coverage) -ForegroundColor Green
