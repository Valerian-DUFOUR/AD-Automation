<#
.SYNOPSIS
    Audit de sécurité des GPO : mots de passe GPP (cpassword), scripts référencés
    et GPO à risque, avec page KPI. Export Excel (.xlsx) sur le Bureau.

.DESCRIPTION
    Parcourt SYSVOL et les GPO du domaine pour repérer :
        - les mots de passe stockés dans les préférences GPP (attribut cpassword,
          déchiffrable publiquement — CRITIQUE, MS14-025) ;
        - les fichiers de scripts (ouverture/fermeture de session, démarrage/arrêt) ;
        - les GPO non liées ou désactivées (bruit / à nettoyer).
    Dates au format JJ-MM-AAAA.

.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-GPOSecurityAudit.ps1

.NOTES
    Prérequis : RSAT ActiveDirectory + GroupPolicy + ImportExcel (auto-installé) ;
    accès en lecture à SYSVOL. Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_GPO_Securite_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
)
$ErrorActionPreference = 'Stop'
$Author = 'Valérian DUFOUR / Claude'
function Write-Step { param($m) Write-Host "[+] $m" -ForegroundColor Cyan }
foreach ($m in 'ActiveDirectory','GroupPolicy') { if (-not (Get-Module -ListAvailable -Name $m)) { throw "Module $m introuvable (RSAT)." }; Import-Module $m -ErrorAction Stop }
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Step "Installation d'ImportExcel..."; try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop } catch { throw "Install-Module ImportExcel -Scope CurrentUser" }
}
Import-Module ImportExcel -ErrorAction Stop

$domain = Get-ADDomain
$sysvol = "\\{0}\SYSVOL\{0}\Policies" -f $domain.DNSRoot
Write-Step "SYSVOL : $sysvol"

# Table GUID -> nom de GPO
$gpoName = @{}
try { Get-GPO -All | ForEach-Object { $gpoName["{$($_.Id)}".ToUpper()] = $_.DisplayName } } catch {}

function Resolve-GpoName { param($path)
    if ($path -match '(\{[0-9A-Fa-f\-]+\})') { $g = $Matches[1].ToUpper(); if ($gpoName.ContainsKey($g)) { return $gpoName[$g] } return $g }
    return '(inconnu)'
}

# 1) cpassword dans les préférences GPP
Write-Step "Recherche des mots de passe GPP (cpassword)..."
$cpwd = New-Object System.Collections.Generic.List[object]
try {
    $xmls = Get-ChildItem -Path $sysvol -Recurse -File -Include *.xml -ErrorAction SilentlyContinue
    foreach ($f in $xmls) {
        $hits = Select-String -Path $f.FullName -Pattern 'cpassword="([^"]+)"' -AllMatches -ErrorAction SilentlyContinue
        foreach ($h in $hits) { foreach ($m in $h.Matches) {
            $val = $m.Groups[1].Value
            $cpwd.Add([PSCustomObject]@{
                'GPO'      = (Resolve-GpoName $f.FullName)
                'Fichier'  = $f.Name
                'Chemin'   = $f.FullName
                'cpassword'= if ($val.Length -gt 12) { $val.Substring(0,12) + '…' } else { $val }
                'Risque'   = 'CRITIQUE'
            })
        } }
    }
} catch { Write-Warning "Scan SYSVOL : $($_.Exception.Message)" }

# 2) Fichiers de scripts
Write-Step "Recensement des fichiers de scripts GPO..."
$scripts = New-Object System.Collections.Generic.List[object]
foreach ($scope in 'Machine','User') {
    try {
        Get-ChildItem -Path $sysvol -Recurse -File -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -match "\\$scope\\Scripts\\" } |
            ForEach-Object { $scripts.Add([PSCustomObject]@{ 'GPO'=(Resolve-GpoName $_.FullName); 'Portée'=$(if($scope -eq 'Machine'){'Ordinateur'}else{'Utilisateur'}); 'Fichier'=$_.FullName }) }
    } catch {}
}

# 3) GPO non liées / désactivées
Write-Step "Analyse des GPO non liées / désactivées..."
$gpoIssues = New-Object System.Collections.Generic.List[object]
try {
    foreach ($g in (Get-GPO -All)) {
        $linked = $true
        try { [xml]$rep = Get-GPOReport -Guid $g.Id -ReportType Xml -ErrorAction Stop; $linked = [bool]$rep.GPO.LinksTo } catch {}
        $issue = @()
        if (-not $linked) { $issue += 'Non liée' }
        if ("$($g.GpoStatus)" -eq 'AllSettingsDisabled') { $issue += 'Toutes désactivées' }
        if ($issue.Count -gt 0) {
            $gpoIssues.Add([PSCustomObject]@{ 'GPO'=$g.DisplayName; 'Problème'=($issue -join ' ; '); 'Date de création'=$g.CreationTime; 'Statut'="$($g.GpoStatus)" })
        }
    }
} catch {}

$nbCpwd = @($cpwd).Count; $nbScripts = @($scripts).Count; $nbIssues = @($gpoIssues).Count
Write-Step ("cpassword: {0} | scripts: {1} | GPO à nettoyer: {2}." -f $nbCpwd,$nbScripts,$nbIssues)

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$script:excel = $null
function Add-Detail { param($Data,$Sheet,$Table,$DateCols)
    if (@($Data).Count -eq 0) { $Data = ,([PSCustomObject]@{ Information='Aucun élément' }) }
    $script:excel = if ($null -eq $script:excel) { $Data | Export-Excel -Path $OutputPath -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
                    else { $Data | Export-Excel -ExcelPackage $script:excel -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
    if ($DateCols) { $ws=$script:excel.Workbook.Worksheets[$Sheet]; $hr=$ws.Dimension.Start.Row; for($c=1;$c -le $ws.Dimension.End.Column;$c++){ if($DateCols -contains $ws.Cells[$hr,$c].Value){ $l=[OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c); $ws.Cells[("{0}{1}:{0}{2}" -f $l,($hr+1),$ws.Dimension.End.Row)].Style.Numberformat.Format='dd-mm-yyyy hh:mm' } } }
}
Add-Detail -Data $cpwd      -Sheet 'Mots de passe GPP' -Table 'GPPcpassword'
Add-Detail -Data $scripts   -Sheet 'Scripts GPO'       -Table 'ScriptsGPO'
Add-Detail -Data $gpoIssues -Sheet 'GPO à nettoyer'    -Table 'GPOanettoyer' -DateCols @('Date de création')
$excel = $script:excel
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Audit de sécurité des GPO'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$kpis = @(
    @('Mots de passe GPP (cpassword)', $nbCpwd, 'red'),
    @('Fichiers de scripts GPO', $nbScripts, $null),
    @('GPO non liées / désactivées', $nbIssues, 'orange')
)
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    if ($k[2] -eq 'red' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,199,206)); $wsK.Cells["C$r"].Style.Font.Bold=$true }
    elseif ($k[2] -eq 'orange' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,235,156)) } }
if ($nbCpwd -gt 0) { $wsK.Cells["B$($r+2)"].Value = "ACTION : supprimez les cpassword (MS14-025) et changez les comptes concernés."; $wsK.Cells["B$($r+2)"].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed); $wsK.Cells["B$($r+2)"].Style.Font.Bold=$true }
$wsK.Column(2).Width=44; $wsK.Column(3).Width=16
Close-ExcelPackage $excel
Write-Host ("`n[OK] Rapport généré : {0}" -f $OutputPath) -ForegroundColor Green
