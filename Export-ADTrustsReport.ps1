<#
.SYNOPSIS
    Rapport des relations d'approbation (trusts) Active Directory, avec page KPI.
    Export Excel (.xlsx) sur le Bureau.

.DESCRIPTION
    Liste les approbations du domaine : partenaire, sens, type, transitivité,
    authentification sélective et filtrage SID (quarantaine). Signale les
    configurations à risque (filtrage SID désactivé sur une approbation externe,
    approbation bidirectionnelle externe, transitivité). Dates au format JJ-MM-AAAA.

.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-ADTrustsReport.ps1

.NOTES
    Prérequis : RSAT ActiveDirectory + ImportExcel (auto-installé).
    Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Trusts_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
)
$ErrorActionPreference = 'Stop'
$Author = 'Valérian DUFOUR / Claude'
function Write-Step { param($m) Write-Host "[+] $m" -ForegroundColor Cyan }
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) { throw "Module ActiveDirectory introuvable (RSAT)." }
Import-Module ActiveDirectory -ErrorAction Stop
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Step "Installation d'ImportExcel..."; try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop } catch { throw "Install-Module ImportExcel -Scope CurrentUser" }
}
Import-Module ImportExcel -ErrorAction Stop

Write-Step "Lecture des approbations..."
$trusts = @()
try { $trusts = Get-ADTrust -Filter * -Properties Direction,TrustType,IntraForest,SelectiveAuthentication,SIDFilteringQuarantined,SIDFilteringForestAware,DisallowTransitivity,ForestTransitive,whenCreated,Target,Source } catch { Write-Warning "Get-ADTrust : $($_.Exception.Message)" }

$report = foreach ($t in $trusts) {
    $sidFilter = if ($t.SIDFilteringQuarantined -or $t.SIDFilteringForestAware) { 'Actif' } else { 'Désactivé' }
    $transitif = if ($t.DisallowTransitivity) { 'Non' } else { 'Oui' }
    $externe = ("$($t.TrustType)" -match 'External') -or (-not $t.IntraForest)
    $risque = 'OK'
    if ($externe -and $sidFilter -eq 'Désactivé') { $risque = 'ELEVE' }
    elseif ("$($t.Direction)" -match 'Bi' -and $externe) { $risque = 'À surveiller' }
    elseif (-not $t.SelectiveAuthentication -and $externe) { $risque = 'À surveiller' }
    [PSCustomObject]@{
        'Partenaire'          = $t.Target
        'Sens'                = "$($t.Direction)"
        'Type'               = "$($t.TrustType)"
        'Intra-forêt'         = if ($t.IntraForest) { 'Oui' } else { 'Non' }
        'Transitif'           = $transitif
        'Auth sélective'      = if ($t.SelectiveAuthentication) { 'Oui' } else { 'Non' }
        'Filtrage SID'        = $sidFilter
        'Date de création'    = $t.whenCreated
        'Risque'              = $risque
    }
}
$report = @($report) | Sort-Object 'Risque','Partenaire'
$total = @($report).Count
$ext   = @($report | Where-Object { $_.'Intra-forêt' -eq 'Non' }).Count
$noSid = @($report | Where-Object { $_.'Filtrage SID' -eq 'Désactivé' }).Count
$bidir = @($report | Where-Object { $_.Sens -match 'Bi' }).Count
Write-Step ("{0} approbation(s) | {1} externe(s) | {2} sans filtrage SID." -f $total,$ext,$noSid)

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$data = if ($total -gt 0) { $report } else { ,([PSCustomObject]@{ Information = 'Aucune approbation' }) }
$excel = $data | Export-Excel -Path $OutputPath -WorksheetName 'Approbations' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName 'Approbations' -TableStyle 'Medium2' -Title ("Relations d'approbation - {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm')) -TitleBold -TitleSize 14 -PassThru
if ($total -gt 0) {
    $ws = $excel.Workbook.Worksheets['Approbations']; $hr = $ws.Dimension.Start.Row
    for ($c=1;$c -le $ws.Dimension.End.Column;$c++){ if ($ws.Cells[$hr,$c].Value -eq 'Date de création'){ $l=[OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c); $ws.Cells[("{0}{1}:{0}{2}" -f $l,($hr+1),$ws.Dimension.End.Row)].Style.Numberformat.Format='dd-mm-yyyy hh:mm' } }
    $rg = "A{0}:{1}{2}" -f ($hr+1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column), $ws.Dimension.End.Row
    # Colonne I = Risque, G = Filtrage SID
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$I{0}="ELEVE"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::MistyRose) -Bold
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$G{0}="Désactivé"' -f ($hr+1)) -ForegroundColor ([System.Drawing.Color]::DarkRed) -Bold
}
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Relations d''approbation (trusts)'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$kpis = @(
    @('Total approbations', $total, $null),
    @('Externes (hors forêt)', $ext, 'orange'),
    @('Sans filtrage SID', $noSid, 'red'),
    @('Bidirectionnelles', $bidir, $null)
)
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    if ($k[2] -eq 'red' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,199,206)); $wsK.Cells["C$r"].Style.Font.Bold=$true }
    elseif ($k[2] -eq 'orange' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,235,156)) } }
$wsK.Column(2).Width=36; $wsK.Column(3).Width=16
Close-ExcelPackage $excel
Write-Host ("`n[OK] Rapport généré : {0}" -f $OutputPath) -ForegroundColor Green
