<#
.SYNOPSIS
    Rapport de couverture LAPS (mots de passe administrateur local gérés) des
    ordinateurs Active Directory, avec page KPI. Export Excel (.xlsx) sur le Bureau.

.DESCRIPTION
    Détecte, pour chaque ordinateur, si un mot de passe administrateur local est
    géré par LAPS, en s'appuyant sur les attributs d'EXPIRATION (lisibles sans
    droit sur le mot de passe lui-même) :
        - LAPS "legacy"  : ms-Mcs-AdmPwdExpirationTime
        - Windows LAPS   : msLAPS-PasswordExpirationTime
    Le script s'adapte automatiquement aux attributs présents dans le schéma.

    Feuilles : « Synthèse (KPI) » (taux de couverture, expirés) et « Ordinateurs »
    (détail, postes non couverts surlignés). Dates au format JJ-MM-AAAA.

.PARAMETER SearchBase
    OU de départ (défaut : tout le domaine).
.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-LAPSReport.ps1

.NOTES
    Prérequis : RSAT ActiveDirectory + ImportExcel (auto-installé).
    Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_LAPS_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
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

function Get-OUFromDN { param($d) if ($d -match '^CN=.*?,(.*)$') { return $Matches[1] }; return $d }

# Détection des attributs LAPS présents dans le schéma
$schema = (Get-ADRootDSE).schemaNamingContext
function Test-Attr { param($n) try { return ([bool](Get-ADObject -SearchBase $schema -LDAPFilter "(lDAPDisplayName=$n)" -ErrorAction Stop)) } catch { return $false } }
$hasLegacy = Test-Attr 'ms-Mcs-AdmPwdExpirationTime'
$hasNew    = Test-Attr 'msLAPS-PasswordExpirationTime'
Write-Step ("LAPS legacy dans le schéma : {0} | Windows LAPS : {1}" -f $hasLegacy, $hasNew)

$props = @('Name','OperatingSystem','LastLogonDate','whenCreated','DistinguishedName','Enabled')
if ($hasLegacy) { $props += 'ms-Mcs-AdmPwdExpirationTime' }
if ($hasNew)    { $props += 'msLAPS-PasswordExpirationTime' }

$get = @{ Filter = 'OperatingSystem -like "*Windows*"'; Properties = $props }
if ($SearchBase) { $get['SearchBase'] = $SearchBase }
Write-Step "Chargement des ordinateurs..."
$computers = Get-ADComputer @get
$now = Get-Date

$report = foreach ($c in $computers) {
    $exp = $null; $type = ''
    if ($hasLegacy -and $c.'ms-Mcs-AdmPwdExpirationTime') { $exp = [datetime]::FromFileTime([int64]$c.'ms-Mcs-AdmPwdExpirationTime'); $type = 'LAPS legacy' }
    if ($hasNew -and $c.'msLAPS-PasswordExpirationTime')  { $exp = [datetime]::FromFileTime([int64]$c.'msLAPS-PasswordExpirationTime'); $type = 'Windows LAPS' }
    $managed = [bool]$exp
    [PSCustomObject]@{
        'Nom du PC'          = $c.Name
        'OS'                 = $c.OperatingSystem
        'LAPS géré'          = if ($managed) { 'Oui' } else { 'Non' }
        'Type LAPS'          = $type
        'Expiration MDP'     = $exp
        'MDP expiré'         = if ($managed -and $exp -lt $now) { 'Oui' } else { 'Non' }
        'Dernière connexion' = $c.LastLogonDate
        'Statut'             = if ($c.Enabled) { 'Activé' } else { 'Désactivé' }
        'Emplacement (OU)'   = (Get-OUFromDN $c.DistinguishedName)
    }
}
$report = @($report) | Sort-Object 'LAPS géré','Nom du PC'

$total   = @($report).Count
$managed = @($report | Where-Object { $_.'LAPS géré' -eq 'Oui' }).Count
$notmgd  = $total - $managed
$expired = @($report | Where-Object { $_.'MDP expiré' -eq 'Oui' }).Count
$cov     = if ($total -gt 0) { [math]::Round(($managed/$total)*100,1) } else { 0 }
Write-Step ("{0} PC | {1} géré(s) LAPS ({2}%) | {3} non couvert(s)." -f $total,$managed,$cov,$notmgd)

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$data = if ($total -gt 0) { $report } else { ,([PSCustomObject]@{ Information = 'Aucun ordinateur' }) }
$excel = $data | Export-Excel -Path $OutputPath -WorksheetName 'Ordinateurs' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName 'Ordinateurs' -TableStyle 'Medium2' -Title ("Couverture LAPS - {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm')) -TitleBold -TitleSize 14 -PassThru
if ($total -gt 0) {
    $ws = $excel.Workbook.Worksheets['Ordinateurs']; $hr = $ws.Dimension.Start.Row
    for ($c=1;$c -le $ws.Dimension.End.Column;$c++){ if (@('Expiration MDP','Dernière connexion') -contains $ws.Cells[$hr,$c].Value){ $l=[OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c); $ws.Cells[("{0}{1}:{0}{2}" -f $l,($hr+1),$ws.Dimension.End.Row)].Style.Numberformat.Format='dd-mm-yyyy hh:mm' } }
    $rg = "A{0}:{1}{2}" -f ($hr+1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column), $ws.Dimension.End.Row
    # C = LAPS géré, F = MDP expiré
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$C{0}="Non"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::MistyRose) -Bold
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$F{0}="Oui"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::Moccasin)
}

$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Couverture LAPS (mots de passe admin locaux)'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
if (-not $hasLegacy -and -not $hasNew) { $wsK.Cells['B5'].Value = "ATTENTION : aucun attribut LAPS dans le schéma — LAPS ne semble pas déployé." }
$kpis = @(
    @('Total ordinateurs', $total, $null),
    @('Gérés par LAPS', $managed, $null),
    @('NON couverts', $notmgd, 'red'),
    @('Taux de couverture (%)', $cov, $null),
    @('Mots de passe expirés', $expired, 'orange')
)
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    if ($k[2] -eq 'red' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,199,206)); $wsK.Cells["C$r"].Style.Font.Bold=$true }
    elseif ($k[2] -eq 'orange' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,235,156)) } }
# graphique couverture
$gr = $r + 3
$wsK.Cells["B$gr"].Value='État'; $wsK.Cells["C$gr"].Value='Nombre'; $wsK.Cells["B$gr:C$gr"].Style.Font.Bold=$true
$wsK.Cells["B$($gr+1)"].Value='Gérés'; $wsK.Cells["C$($gr+1)"].Value=$managed
$wsK.Cells["B$($gr+2)"].Value='Non couverts'; $wsK.Cells["C$($gr+2)"].Value=$notmgd
try { $ch=$wsK.Drawings.AddChart('lapsChart',[OfficeOpenXml.Drawing.Chart.eChartType]::Pie); $ch.Title.Text='Couverture LAPS'; $ch.SetPosition(6,0,5,0); $ch.SetSize(400,280); $null=$ch.Series.Add($wsK.Cells["C$($gr+1):C$($gr+2)"],$wsK.Cells["B$($gr+1):B$($gr+2)"]); $ch.DataLabel.ShowPercent=$true } catch {}
$wsK.Column(2).Width=42; $wsK.Column(3).Width=16
Close-ExcelPackage $excel
Write-Host ("`n[OK] Rapport généré : {0}" -f $OutputPath) -ForegroundColor Green
