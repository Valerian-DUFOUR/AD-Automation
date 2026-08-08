<#
.SYNOPSIS
    Extraction des stratégies de groupe (GPO) Active Directory et génération
    d'un rapport Excel (.xlsx) formaté sur le Bureau.

.DESCRIPTION
    Ce script interroge les GPO du domaine via le module GroupPolicy
    (Get-GPO / Get-GPOReport) et produit un classeur Excel à plusieurs
    feuilles :

        1. « Synthèse »       : indicateurs clés (KPI) de l'extraction.
        2. « GPO »            : une GPO par ligne, avec sa date de création,
           son statut, sa description, le chemin de son dossier SYSVOL,
           le nombre de liens, le nombre d'entités de filtrage de sécurité
           (« membres » à qui la GPO s'applique), et le filtre WMI.
        3. « Liens »          : détail des liens (une ligne par OU/site lié).
        4. « Fichiers scripts » : chemins des fichiers de scripts référencés
           par les GPO (ouverture/fermeture de session, démarrage/arrêt),
           s'il y en a.

    Remarque sur « le nombre de membres par GPO » : une GPO n'a pas de
    membres au sens strict. Ce script fournit deux mesures : le nombre de
    LIENS (OU/sites où la GPO est appliquée) et le nombre d'ENTITÉS DE
    FILTRAGE DE SÉCURITÉ (comptes/groupes ayant le droit « Appliquer la
    stratégie de groupe »), qui déterminent à qui la GPO s'applique.

    Toutes les dates sont au format JJ-MM-AAAA.

.PARAMETER OutputPath
    (Optionnel) Chemin complet du fichier .xlsx de sortie.
    Par défaut : le Bureau de l'utilisateur courant.

.EXAMPLE
    .\Export-GPOReport.ps1

.NOTES
    Prérequis :
        - Windows avec le module RSAT « Gestion des stratégies de groupe »
          (module GroupPolicy : Get-Command Get-GPO doit fonctionner)
        - Module ImportExcel (installé automatiquement si absent)
        - Droits de lecture sur les GPO et, idéalement, sur SYSVOL
          (pour lister les fichiers de scripts)

    Auteur : Valérian DUFOUR / Claude
#>

[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_GPO_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
)

$ErrorActionPreference = 'Stop'
$Author = 'Valérian DUFOUR / Claude'

function Write-Step { param([string]$Message) Write-Host "[+] $Message" -ForegroundColor Cyan }

# --------------------------------------------------------------------------
# 1. Prérequis
# --------------------------------------------------------------------------
Write-Step "Vérification du module GroupPolicy..."
if (-not (Get-Module -ListAvailable -Name GroupPolicy)) {
    throw "Le module 'GroupPolicy' est introuvable. Installez RSAT « Gestion des stratégies de groupe » puis relancez."
}
Import-Module GroupPolicy -ErrorAction Stop

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
function Translate-Status {
    param($s)
    switch ("$s") {
        'AllSettingsEnabled'       { 'Toutes activées' }
        'AllSettingsDisabled'      { 'Toutes désactivées' }
        'UserSettingsDisabled'     { 'Config. utilisateur désactivée' }
        'ComputerSettingsDisabled' { 'Config. ordinateur désactivée' }
        default                    { "$s" }
    }
}

# --------------------------------------------------------------------------
# 3. Récupération des GPO
# --------------------------------------------------------------------------
Write-Step "Interrogation des GPO du domaine..."
$gpos = Get-GPO -All | Sort-Object DisplayName
Write-Step ("{0} GPO récupérée(s)." -f $gpos.Count)

$now = Get-Date
$gpoRows   = [System.Collections.Generic.List[object]]::new()
$linkRows  = [System.Collections.Generic.List[object]]::new()
$fileRows  = [System.Collections.Generic.List[object]]::new()

$total = $gpos.Count
$idx = 0
foreach ($gpo in $gpos) {
    $idx++
    Write-Progress -Activity "Analyse des GPO" -Status "$($gpo.DisplayName) ($idx/$total)" `
                   -PercentComplete ([int](($idx / [Math]::Max($total,1)) * 100))

    # -- Chemin SYSVOL de la GPO --
    $sysvolPath = "\\{0}\SYSVOL\{0}\Policies\{{{1}}}" -f $gpo.DomainName, $gpo.Id

    # -- Liens (via le rapport XML) --
    $links = @()
    try {
        [xml]$rep = Get-GPOReport -Guid $gpo.Id -ReportType Xml -ErrorAction Stop
        if ($rep.GPO.LinksTo) { $links = @($rep.GPO.LinksTo) }
    } catch {
        Write-Warning "Rapport XML indisponible pour '$($gpo.DisplayName)' : $($_.Exception.Message)"
    }
    foreach ($l in $links) {
        $linkRows.Add([PSCustomObject]@{
            'GPO'      = $gpo.DisplayName
            'Lien (emplacement)' = $l.SOMPath
            'Activé'   = if ("$($l.Enabled)" -eq 'true') { 'Oui' } else { 'Non' }
            'Appliqué (No Override)' = if ("$($l.NoOverride)" -eq 'true') { 'Oui' } else { 'Non' }
        })
    }

    # -- Filtrage de sécurité (entités avec « Appliquer la stratégie de groupe ») --
    $filterNames = @()
    try {
        $filterNames = @(Get-GPPermission -Guid $gpo.Id -All -ErrorAction Stop |
            Where-Object { $_.Permission -eq 'GpoApply' } |
            ForEach-Object { $_.Trustee.Name })
    } catch { }

    # -- Fichiers de scripts présents dans SYSVOL --
    $scriptFiles = @()
    foreach ($scope in @('Machine','User')) {
        $scriptsDir = Join-Path $sysvolPath "$scope\Scripts"
        try {
            if (Test-Path $scriptsDir) {
                Get-ChildItem -Path $scriptsDir -Recurse -File -ErrorAction SilentlyContinue |
                    ForEach-Object {
                        $scriptFiles += $_.FullName
                        $fileRows.Add([PSCustomObject]@{
                            'GPO'     = $gpo.DisplayName
                            'Portée'  = if ($scope -eq 'Machine') { 'Ordinateur' } else { 'Utilisateur' }
                            'Fichier' = $_.FullName
                        })
                    }
            }
        } catch { }
    }

    $gpoRows.Add([PSCustomObject]@{
        'Nom de la GPO'          = $gpo.DisplayName
        'Statut'                 = (Translate-Status $gpo.GpoStatus)
        'Description'            = $gpo.Description
        'Date de création'       = $gpo.CreationTime
        'Date de modification'   = $gpo.ModificationTime
        'Propriétaire'           = $gpo.Owner
        'Filtre WMI'             = if ($gpo.WmiFilter) { $gpo.WmiFilter.Name } else { '' }
        'Nb liens'               = @($links).Count
        'Liens (emplacements)'   = (@($links | ForEach-Object { $_.SOMPath }) -join ' ; ')
        'Nb entités (filtrage)'  = @($filterNames).Count
        'Filtrage de sécurité'   = ($filterNames -join ' ; ')
        'Nb fichiers scripts'    = @($scriptFiles).Count
        'Chemin SYSVOL'          = $sysvolPath
        'GUID'                   = "{$($gpo.Id)}"
    })
}
Write-Progress -Activity "Analyse des GPO" -Completed

$gpoReport  = $gpoRows
$linkReport = $linkRows | Sort-Object 'GPO'
$fileReport = $fileRows | Sort-Object 'GPO'

# --------------------------------------------------------------------------
# 4. KPI
# --------------------------------------------------------------------------
$totalGpo    = @($gpoReport).Count
$enabled     = @($gpoReport | Where-Object { $_.Statut -eq 'Toutes activées' }).Count
$disabled    = @($gpoReport | Where-Object { $_.Statut -eq 'Toutes désactivées' }).Count
$unlinked    = @($gpoReport | Where-Object { $_.'Nb liens' -eq 0 }).Count
$withScripts = @($gpoReport | Where-Object { $_.'Nb fichiers scripts' -gt 0 }).Count

$pct = { param($n) if ($totalGpo -gt 0) { [math]::Round(($n / $totalGpo) * 100, 1) } else { 0 } }

$statusDist = $gpoReport | Group-Object 'Statut' | Sort-Object Count -Descending |
    ForEach-Object { [PSCustomObject]@{ 'Statut' = $_.Name; 'Nombre' = $_.Count } }

# --------------------------------------------------------------------------
# 5. Export Excel
# --------------------------------------------------------------------------
Write-Step "Génération du fichier Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

$dateFmt = 'dd-mm-yyyy hh:mm'
function Format-DateColumns {
    param($ws, [string[]]$Columns)
    $hr = $ws.Dimension.Start.Row
    for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
        if ($Columns -contains $ws.Cells[$hr, $c].Value) {
            $l = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c)
            $ws.Cells[("{0}{1}:{0}{2}" -f $l, ($hr + 1), $ws.Dimension.End.Row)].Style.Numberformat.Format = $dateFmt
        }
    }
}

# --- Feuille GPO ---
$excel = $gpoReport | Export-Excel -Path $OutputPath -WorksheetName 'GPO' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'GPO' -TableStyle 'Medium2' `
    -Title ("Stratégies de groupe (GPO) - {0}" -f $now.ToString('dd-MM-yyyy HH:mm')) -TitleBold -TitleSize 14 -PassThru

$wsG = $excel.Workbook.Worksheets['GPO']
Format-DateColumns -ws $wsG -Columns @('Date de création','Date de modification')
$hrG = $wsG.Dimension.Start.Row
$rangeG = "A{0}:{1}{2}" -f ($hrG + 1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($wsG.Dimension.End.Column), $wsG.Dimension.End.Row
# Colonne B = Statut ; H = Nb liens
Add-ConditionalFormatting -Worksheet $wsG -Range $rangeG -RuleType Expression `
    -ConditionValue ('=$B{0}="Toutes désactivées"' -f ($hrG + 1)) -ForegroundColor ([System.Drawing.Color]::Gray) -Italic
Add-ConditionalFormatting -Worksheet $wsG -Range $rangeG -RuleType Expression `
    -ConditionValue ('=$H{0}=0' -f ($hrG + 1)) -BackgroundColor ([System.Drawing.Color]::MistyRose)

# --- Feuille Liens ---
if (@($linkReport).Count -gt 0) {
    $excel = $linkReport | Export-Excel -ExcelPackage $excel -WorksheetName 'Liens' `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'Liens' -TableStyle 'Medium2' -PassThru
}

# --- Feuille Fichiers scripts ---
if (@($fileReport).Count -gt 0) {
    $excel = $fileReport | Export-Excel -ExcelPackage $excel -WorksheetName 'Fichiers scripts' `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'FichiersScripts' -TableStyle 'Medium2' -PassThru
}

# --- Feuille Synthèse (1ère position) ---
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse' -MoveToStart

$wsK.Cells['B2'].Value = 'Rapport des stratégies de groupe (GPO)'
$wsK.Cells['B2'].Style.Font.Size = 16
$wsK.Cells['B2'].Style.Font.Bold = $true
$wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f $now.ToString('dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)

$kpis = @(
    @('Total des GPO',                 $totalGpo,    $null),
    @('GPO toutes activées',           $enabled,     (& $pct $enabled)),
    @('GPO toutes désactivées',        $disabled,    (& $pct $disabled)),
    @('GPO non liées (0 lien)',        $unlinked,    (& $pct $unlinked)),
    @('GPO avec fichiers scripts',     $withScripts, (& $pct $withScripts))
)

$row = 6
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
$kpiLastRow = $row

# Répartition par statut
$stHeaderRow = $kpiLastRow + 3
$wsK.Cells["B$stHeaderRow"].Value = 'Statut'; $wsK.Cells["C$stHeaderRow"].Value = 'Nombre'
$wsK.Cells["B$stHeaderRow:C$stHeaderRow"].Style.Font.Bold = $true
$wsK.Cells["B$stHeaderRow:C$stHeaderRow"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$stHeaderRow:C$stHeaderRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$stHeaderRow:C$stHeaderRow"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
$r = $stHeaderRow
foreach ($s in $statusDist) { $r++; $wsK.Cells["B$r"].Value = $s.'Statut'; $wsK.Cells["C$r"].Value = $s.'Nombre' }
$stLastRow = $r

try {
    $chart = $wsK.Drawings.AddChart('gpoChart', [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
    $chart.Title.Text = 'Répartition des GPO par statut'
    $chart.SetPosition(5, 0, 5, 0); $chart.SetSize(460, 300)
    $serie = $chart.Series.Add($wsK.Cells["C$($stHeaderRow + 1):C$stLastRow"], $wsK.Cells["B$($stHeaderRow + 1):B$stLastRow"])
    $chart.DataLabel.ShowPercent = $true
} catch { Write-Warning "Graphique non généré : $($_.Exception.Message)" }

$wsK.Column(2).Width = 34; $wsK.Column(3).Width = 14; $wsK.Column(4).Width = 10

Close-ExcelPackage $excel

Write-Host ""
Write-Host "[OK] Rapport généré : $OutputPath" -ForegroundColor Green
Write-Host ("     {0} GPO | {1} lien(s) | {2} fichier(s) de scripts." -f `
    $totalGpo, @($linkReport).Count, @($fileReport).Count) -ForegroundColor Green
