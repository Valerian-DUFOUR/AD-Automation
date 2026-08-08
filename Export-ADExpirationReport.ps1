<#
.SYNOPSIS
    Rapport des échéances et objets à nettoyer dans Active Directory : comptes et
    mots de passe expirant bientôt, comptes désactivés, groupes et OU vides, avec
    page KPI. Export Excel (.xlsx) sur le Bureau.

.DESCRIPTION
    Regroupe plusieurs vues d'hygiène AD (dates au format JJ-MM-AAAA) :
        - comptes dont le compte expire dans les N prochains jours (ou déjà expiré) ;
        - comptes dont le mot de passe expire dans les N prochains jours ;
        - comptes désactivés (jamais supprimés) ;
        - groupes sans membre ;
        - OU sans objet.

.PARAMETER Days
    Fenêtre d'échéance en jours (défaut : 30).
.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-ADExpirationReport.ps1 -Days 45

.NOTES
    Prérequis : RSAT ActiveDirectory + ImportExcel (auto-installé).
    Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [int]$Days = 30,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Echeances_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
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

$now = Get-Date; $limit = $now.AddDays($Days)
Write-Step "Chargement des utilisateurs..."
$users = Get-ADUser -Filter * -Properties DisplayName,AccountExpirationDate,'msDS-UserPasswordExpiryTimeComputed',PasswordNeverExpires,Enabled,LastLogonDate,whenCreated,DistinguishedName

$acctExpiring = $users | Where-Object { $_.AccountExpirationDate -and $_.AccountExpirationDate -le $limit } | ForEach-Object {
    [PSCustomObject]@{ 'Nom du compte'=$_.SamAccountName; 'Nom complet'=$_.DisplayName; 'Expiration compte'=$_.AccountExpirationDate; 'Statut'=(if($_.Enabled){'Activé'}else{'Désactivé'}); 'Expiré'=(if($_.AccountExpirationDate -lt $now){'Oui'}else{'Non'}); 'Emplacement (OU)'=(Get-OUFromDN $_.DistinguishedName) }
} | Sort-Object 'Expiration compte'

$pwdExpiring = $users | Where-Object { $_.Enabled -and -not $_.PasswordNeverExpires -and $_.'msDS-UserPasswordExpiryTimeComputed' -and $_.'msDS-UserPasswordExpiryTimeComputed' -ne 0 -and $_.'msDS-UserPasswordExpiryTimeComputed' -ne 9223372036854775807 } | ForEach-Object {
    $exp = [datetime]::FromFileTime([int64]$_.'msDS-UserPasswordExpiryTimeComputed')
    [PSCustomObject]@{ u=$_; exp=$exp }
} | Where-Object { $_.exp -le $limit } | ForEach-Object {
    [PSCustomObject]@{ 'Nom du compte'=$_.u.SamAccountName; 'Nom complet'=$_.u.DisplayName; 'Expiration MDP'=$_.exp; 'Expiré'=(if($_.exp -lt $now){'Oui'}else{'Non'}); 'Emplacement (OU)'=(Get-OUFromDN $_.u.DistinguishedName) }
} | Sort-Object 'Expiration MDP'

$disabled = $users | Where-Object { -not $_.Enabled } | ForEach-Object {
    [PSCustomObject]@{ 'Nom du compte'=$_.SamAccountName; 'Nom complet'=$_.DisplayName; 'Dernière connexion'=$_.LastLogonDate; 'Date de création'=$_.whenCreated; 'Emplacement (OU)'=(Get-OUFromDN $_.DistinguishedName) }
} | Sort-Object 'Nom du compte'

Write-Step "Analyse des groupes vides..."
$emptyGroups = Get-ADGroup -Filter * -Properties Members,whenCreated | Where-Object { @($_.Members).Count -eq 0 } | ForEach-Object {
    [PSCustomObject]@{ 'Groupe'=$_.Name; 'Date de création'=$_.whenCreated; 'Emplacement (OU)'=(Get-OUFromDN $_.DistinguishedName) }
} | Sort-Object 'Groupe'

Write-Step "Analyse des OU vides..."
$emptyOUs = foreach ($ou in (Get-ADOrganizationalUnit -Filter * -ResultSetSize $null)) {
    $child = @(Get-ADObject -SearchBase $ou.DistinguishedName -SearchScope OneLevel -Filter * -ErrorAction SilentlyContinue | Where-Object { $_.DistinguishedName -ne $ou.DistinguishedName })
    if ($child.Count -eq 0) { [PSCustomObject]@{ 'OU'=$ou.Name; 'DN'=$ou.DistinguishedName } }
}
$emptyOUs = @($emptyOUs) | Sort-Object 'DN'

$kAcct=@($acctExpiring).Count; $kPwd=@($pwdExpiring).Count; $kDis=@($disabled).Count; $kGrp=@($emptyGroups).Count; $kOu=@($emptyOUs).Count
Write-Step ("Comptes expirant:{0} MDP:{1} Désactivés:{2} Groupes vides:{3} OU vides:{4}" -f $kAcct,$kPwd,$kDis,$kGrp,$kOu)

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$script:excel = $null
function Add-Detail { param($Data,$Sheet,$Table,$DateCols)
    if (@($Data).Count -eq 0) { $Data = ,([PSCustomObject]@{ Information='Aucun élément' }) }
    $script:excel = if ($null -eq $script:excel) { $Data | Export-Excel -Path $OutputPath -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
                    else { $Data | Export-Excel -ExcelPackage $script:excel -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
    if ($DateCols) { $ws=$script:excel.Workbook.Worksheets[$Sheet]; $hr=$ws.Dimension.Start.Row; for($c=1;$c -le $ws.Dimension.End.Column;$c++){ if($DateCols -contains $ws.Cells[$hr,$c].Value){ $l=[OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c); $ws.Cells[("{0}{1}:{0}{2}" -f $l,($hr+1),$ws.Dimension.End.Row)].Style.Numberformat.Format='dd-mm-yyyy hh:mm' } } }
}
Add-Detail -Data $acctExpiring -Sheet 'Comptes expirant'   -Table 'ComptesExpirant' -DateCols @('Expiration compte')
Add-Detail -Data $pwdExpiring  -Sheet 'MDP expirant'        -Table 'MDPExpirant'    -DateCols @('Expiration MDP')
Add-Detail -Data $disabled     -Sheet 'Comptes désactivés'  -Table 'Desactives'     -DateCols @('Dernière connexion','Date de création')
Add-Detail -Data $emptyGroups  -Sheet 'Groupes vides'       -Table 'GroupesVides'   -DateCols @('Date de création')
Add-Detail -Data $emptyOUs     -Sheet 'OU vides'            -Table 'OUVides'
$excel = $script:excel

$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Échéances & nettoyage Active Directory'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}  |  Fenêtre : {1} jours" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'), $Days)
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$kpis = @(
    @("Comptes expirant (< $Days j)", $kAcct, 'orange'),
    @("Mots de passe expirant (< $Days j)", $kPwd, 'orange'),
    @('Comptes désactivés', $kDis, $null),
    @('Groupes vides', $kGrp, $null),
    @('OU vides', $kOu, $null)
)
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    if ($k[2] -eq 'orange' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,235,156)) }
    elseif ($r % 2 -eq 0) { $wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217,226,243)) } }
$wsK.Column(2).Width=42; $wsK.Column(3).Width=16
Close-ExcelPackage $excel
Write-Host ("`n[OK] Rapport généré : {0}" -f $OutputPath) -ForegroundColor Green
