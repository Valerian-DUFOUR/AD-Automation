<#
.SYNOPSIS
    Rapport de santé Active Directory : rôles FSMO, contrôleurs de domaine,
    réplication, avec page KPI. Export Excel (.xlsx) sur le Bureau.

.DESCRIPTION
    Synthétise l'état de l'annuaire : titulaires des rôles FSMO, joignabilité des
    DC (ping + LDAP 389), et échecs de réplication (Get-ADReplicationFailure).
    Dates au format JJ-MM-AAAA.

.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-ADHealthReport.ps1

.NOTES
    Prérequis : RSAT ActiveDirectory + ImportExcel (auto-installé). À lancer près
    des DC. Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Sante_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
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

$forest = Get-ADForest; $domain = Get-ADDomain
$fsmo = @(
    [PSCustomObject]@{ 'Rôle'='Schema Master';         'Titulaire'=$forest.SchemaMaster }
    [PSCustomObject]@{ 'Rôle'='Domain Naming Master';  'Titulaire'=$forest.DomainNamingMaster }
    [PSCustomObject]@{ 'Rôle'='PDC Emulator';          'Titulaire'=$domain.PDCEmulator }
    [PSCustomObject]@{ 'Rôle'='RID Master';            'Titulaire'=$domain.RIDMaster }
    [PSCustomObject]@{ 'Rôle'='Infrastructure Master'; 'Titulaire'=$domain.InfrastructureMaster }
)

Write-Step "Analyse des contrôleurs de domaine..."
$dcs = Get-ADDomainController -Filter *
$dcReport = foreach ($dc in $dcs) {
    $ping = $false; try { $ping = Test-Connection -ComputerName $dc.HostName -Count 1 -Quiet -ErrorAction SilentlyContinue } catch {}
    $ldap = $false; try { $t = New-Object System.Net.Sockets.TcpClient; $a=$t.BeginConnect($dc.HostName,389,$null,$null); if ($a.AsyncWaitHandle.WaitOne(1500,$false) -and $t.Connected){$ldap=$true}; $t.Close() } catch {}
    $fails = 0; try { $fails = @(Get-ADReplicationFailure -Target $dc.HostName -ErrorAction Stop).Count } catch { $fails = -1 }
    [PSCustomObject]@{
        'Contrôleur'        = $dc.HostName
        'Site'              = $dc.Site
        'OS'                = $dc.OperatingSystem
        'Global Catalog'    = if ($dc.IsGlobalCatalog) { 'Oui' } else { 'Non' }
        'Ping'              = if ($ping) { 'OK' } else { 'KO' }
        'LDAP 389'          = if ($ldap) { 'OK' } else { 'KO' }
        'Échecs réplication'= if ($fails -lt 0) { 'N/A' } else { $fails }
        'État'              = if (-not $ping -or -not $ldap) { 'INJOIGNABLE' } elseif ($fails -gt 0) { 'RÉPLICATION KO' } else { 'OK' }
    }
}

# Métadonnées de réplication (dernier succès par partenaire)
$repl = @()
try {
    foreach ($dc in $dcs) {
        try { Get-ADReplicationPartnerMetadata -Target $dc.HostName -ErrorAction Stop | ForEach-Object {
            $repl += [PSCustomObject]@{ 'DC'=$dc.HostName; 'Partenaire'=($_.Partner -replace '^CN=NTDS Settings,CN=',''); 'Dernier succès'=$_.LastReplicationSuccess; 'Dernier échec'=$_.LastReplicationResult; 'Nb échecs consécutifs'=$_.ConsecutiveReplicationFailures }
        } } catch {}
    }
} catch {}

$nbDC = @($dcReport).Count
$nbKO = @($dcReport | Where-Object { $_.'État' -eq 'INJOIGNABLE' }).Count
$nbReplKO = @($dcReport | Where-Object { $_.'État' -eq 'RÉPLICATION KO' }).Count
Write-Step ("{0} DC | {1} injoignable(s) | {2} avec échecs réplication." -f $nbDC,$nbKO,$nbReplKO)

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$script:excel = $null
function Add-Detail { param($Data,$Sheet,$Table,$DateCols)
    if (@($Data).Count -eq 0) { $Data = ,([PSCustomObject]@{ Information='Aucun élément' }) }
    $script:excel = if ($null -eq $script:excel) { $Data | Export-Excel -Path $OutputPath -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
                    else { $Data | Export-Excel -ExcelPackage $script:excel -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
    if ($DateCols) { $ws=$script:excel.Workbook.Worksheets[$Sheet]; $hr=$ws.Dimension.Start.Row; for($c=1;$c -le $ws.Dimension.End.Column;$c++){ if($DateCols -contains $ws.Cells[$hr,$c].Value){ $l=[OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c); $ws.Cells[("{0}{1}:{0}{2}" -f $l,($hr+1),$ws.Dimension.End.Row)].Style.Numberformat.Format='dd-mm-yyyy hh:mm' } } }
}
Add-Detail -Data $dcReport -Sheet 'Contrôleurs' -Table 'DC'
if (@($dcReport).Count -gt 0) {
    $ws=$script:excel.Workbook.Worksheets['Contrôleurs']; $hr=$ws.Dimension.Start.Row
    $rg = "A{0}:{1}{2}" -f ($hr+1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column), $ws.Dimension.End.Row
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$H{0}="INJOIGNABLE"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::MistyRose) -Bold
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$H{0}="RÉPLICATION KO"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::Moccasin)
}
Add-Detail -Data $fsmo -Sheet 'FSMO' -Table 'FSMO'
Add-Detail -Data $repl -Sheet 'Réplication' -Table 'Replication' -DateCols @('Dernier succès')
$excel = $script:excel

$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Santé Active Directory'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}  |  Forêt : {1}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'), $forest.Name)
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$kpis = @(
    @('Contrôleurs de domaine', $nbDC, $null),
    @('DC injoignables', $nbKO, 'red'),
    @('DC avec échecs de réplication', $nbReplKO, 'red'),
    @('Niveau fonctionnel forêt', "$($forest.ForestMode)", $null),
    @('Niveau fonctionnel domaine', "$($domain.DomainMode)", $null)
)
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    if ($k[2] -eq 'red' -and ($k[1] -is [int]) -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,199,206)); $wsK.Cells["C$r"].Style.Font.Bold=$true } }
$wsK.Column(2).Width=40; $wsK.Column(3).Width=24
Close-ExcelPackage $excel
Write-Host ("`n[OK] Rapport généré : {0}" -f $OutputPath) -ForegroundColor Green
