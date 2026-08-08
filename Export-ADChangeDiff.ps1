<#
.SYNOPSIS
    Compare l'état d'Active Directory entre deux instantanés (snapshots) pour
    détecter les dérives : nouveaux comptes, comptes supprimés/réactivés, et
    surtout les changements d'appartenance aux groupes privilégiés. Page KPI.

.DESCRIPTION
    À chaque exécution, le script crée un instantané (JSON daté) de l'AD
    (utilisateurs, ordinateurs, membres des groupes privilégiés) dans un dossier
    dédié, puis le compare au précédent instantané disponible et génère un
    rapport Excel des changements. Le premier passage crée simplement la
    référence. Dates au format JJ-MM-AAAA.

.PARAMETER SnapshotDir
    Dossier des instantanés (défaut : Bureau\AD_Snapshots).
.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-ADChangeDiff.ps1        # à planifier quotidiennement

.NOTES
    Prérequis : RSAT ActiveDirectory + ImportExcel (auto-installé).
    Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string]$SnapshotDir = (Join-Path ([Environment]::GetFolderPath('Desktop')) 'AD_Snapshots'),
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Changements_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
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
if (-not (Test-Path $SnapshotDir)) { New-Item -ItemType Directory -Path $SnapshotDir -Force | Out-Null }

# --- Construction de l'instantané courant ---
Write-Step "Construction de l'instantané..."
$domain = Get-ADDomain
$usersH = @{}; Get-ADUser -Filter * -Properties Enabled,DisplayName | ForEach-Object { $usersH[$_.SamAccountName] = [bool]$_.Enabled }
$compsH = @{}; Get-ADComputer -Filter * -Properties Enabled | ForEach-Object { $compsH[$_.SamAccountName] = [bool]$_.Enabled }
$priv = @{}
foreach ($rid in 512,519,518,544) {
    $sid = if ($rid -eq 544) { 'S-1-5-32-544' } else { "$($domain.DomainSID.Value)-$rid" }
    try { $g = Get-ADGroup -Identity $sid -ErrorAction Stop
        $members = @(Get-ADGroupMember -Identity $g -Recursive -ErrorAction SilentlyContinue | Where-Object { $_.objectClass -eq 'user' } | Select-Object -ExpandProperty SamAccountName)
        $priv[$g.Name] = $members
    } catch {}
}
$snapshot = [PSCustomObject]@{ Date = (Get-Date).ToString('o'); Users = $usersH; Computers = $compsH; Priv = $priv }
$snapFile = Join-Path $SnapshotDir ("Snapshot_{0}.json" -f (Get-Date -Format 'yyyyMMdd_HHmmss'))
$snapshot | ConvertTo-Json -Depth 6 | Set-Content -Path $snapFile -Encoding UTF8
Write-Step "Instantané enregistré : $snapFile"

# --- Recherche de l'instantané précédent ---
$prevFile = Get-ChildItem -Path $SnapshotDir -Filter 'Snapshot_*.json' | Where-Object { $_.FullName -ne $snapFile } | Sort-Object LastWriteTime -Descending | Select-Object -First 1
$newUsers=@(); $delUsers=@(); $reUsers=@(); $newComps=@(); $delComps=@(); $newAdmins=@(); $delAdmins=@()
$baseline = $false
if (-not $prevFile) {
    $baseline = $true
    Write-Step "Aucun instantané précédent : référence créée (pas de comparaison)."
} else {
    Write-Step "Comparaison avec : $($prevFile.Name)"
    $old = Get-Content $prevFile.FullName -Raw | ConvertFrom-Json
    $oldU = @{}; $old.Users.PSObject.Properties | ForEach-Object { $oldU[$_.Name] = $_.Value }
    $oldC = @{}; $old.Computers.PSObject.Properties | ForEach-Object { $oldC[$_.Name] = $_.Value }
    foreach ($s in $usersH.Keys) {
        if (-not $oldU.ContainsKey($s)) { $newUsers += [PSCustomObject]@{ 'Compte'=$s; 'Statut'=(if($usersH[$s]){'Activé'}else{'Désactivé'}) } }
        elseif ((-not $oldU[$s]) -and $usersH[$s]) { $reUsers += [PSCustomObject]@{ 'Compte'=$s; 'Changement'='Réactivé (Désactivé -> Activé)' } }
    }
    foreach ($s in $oldU.Keys) { if (-not $usersH.ContainsKey($s)) { $delUsers += [PSCustomObject]@{ 'Compte'=$s } } }
    foreach ($s in $compsH.Keys) { if (-not $oldC.ContainsKey($s)) { $newComps += [PSCustomObject]@{ 'Ordinateur'=$s } } }
    foreach ($s in $oldC.Keys) { if (-not $compsH.ContainsKey($s)) { $delComps += [PSCustomObject]@{ 'Ordinateur'=$s } } }
    $oldP = @{}; $old.Priv.PSObject.Properties | ForEach-Object { $oldP[$_.Name] = @($_.Value) }
    foreach ($g in $priv.Keys) {
        $before = @(); if ($oldP.ContainsKey($g)) { $before = $oldP[$g] }
        foreach ($m in $priv[$g]) { if ($before -notcontains $m) { $newAdmins += [PSCustomObject]@{ 'Groupe'=$g; 'Compte'=$m; 'Changement'='AJOUTÉ' } } }
        foreach ($m in $before) { if ($priv[$g] -notcontains $m) { $delAdmins += [PSCustomObject]@{ 'Groupe'=$g; 'Compte'=$m; 'Changement'='RETIRÉ' } } }
    }
}

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$script:excel = $null
function Add-Detail { param($Data,$Sheet,$Table)
    if (@($Data).Count -eq 0) { $Data = ,([PSCustomObject]@{ Information=(if($baseline){'Référence initiale - relancez plus tard pour comparer'}else{'Aucun changement'}) }) }
    $script:excel = if ($null -eq $script:excel) { $Data | Export-Excel -Path $OutputPath -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
                    else { $Data | Export-Excel -ExcelPackage $script:excel -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
}
Add-Detail -Data $newAdmins -Sheet 'Admins ajoutés'    -Table 'AdminsAjoutes'
Add-Detail -Data $delAdmins -Sheet 'Admins retirés'    -Table 'AdminsRetires'
Add-Detail -Data $newUsers  -Sheet 'Nouveaux comptes'  -Table 'NouveauxComptes'
Add-Detail -Data $delUsers  -Sheet 'Comptes supprimés' -Table 'ComptesSupprimes'
Add-Detail -Data $reUsers   -Sheet 'Comptes réactivés' -Table 'ComptesReactives'
Add-Detail -Data $newComps  -Sheet 'Nouveaux PC'       -Table 'NouveauxPC'
Add-Detail -Data $delComps  -Sheet 'PC supprimés'      -Table 'PCSupprimes'
$excel = $script:excel

if (@($newAdmins).Count -gt 0) {
    $ws=$excel.Workbook.Worksheets['Admins ajoutés']; $hr=$ws.Dimension.Start.Row
    $rg = "A{0}:{1}{2}" -f ($hr+1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column), $ws.Dimension.End.Row
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$C{0}="AJOUTÉ"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::MistyRose) -Bold
}

$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Changements Active Directory'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$wsK.Cells['B5'].Value = if ($baseline) { "Référence initiale créée : $($snapFile | Split-Path -Leaf)" } else { "Comparé à : $($prevFile.Name)" }
$kpis = @(
    @('Admins AJOUTÉS', @($newAdmins).Count, 'red'),
    @('Admins RETIRÉS', @($delAdmins).Count, 'orange'),
    @('Nouveaux comptes', @($newUsers).Count, $null),
    @('Comptes supprimés', @($delUsers).Count, $null),
    @('Comptes réactivés', @($reUsers).Count, 'orange'),
    @('Nouveaux ordinateurs', @($newComps).Count, $null),
    @('Ordinateurs supprimés', @($delComps).Count, $null)
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
