<#
.SYNOPSIS
    Rapport des délégations et ACL sensibles dans Active Directory, avec page KPI.
    Export Excel (.xlsx) sur le Bureau.

.DESCRIPTION
    Analyse les listes de contrôle d'accès (ACL) des OU, de la racine du domaine,
    de l'objet AdminSDHolder et des groupes privilégiés, et signale les droits
    dangereux accordés à des comptes/groupes NON administrateurs :
        - GenericAll / GenericWrite / WriteDacl / WriteOwner
        - Réinitialisation de mot de passe (extended right)
        - DCSync (Replicating Directory Changes) — critique
    C'est une vue « BloodHound allégée » des chemins d'escalade de privilèges.
    Dates au format JJ-MM-AAAA.

.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-ADDelegationReport.ps1

.NOTES
    Prérequis : RSAT ActiveDirectory + ImportExcel (auto-installé). Lecture des ACL.
    Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Delegations_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
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

$domain = Get-ADDomain
# GUID des droits étendus / attributs sensibles
$G_RESETPWD = '00299570-246d-11d0-a768-00aa006e0529'
$G_DCSYNC1  = '1131f6aa-9c07-11d1-f79f-00c04fc2dcd2'
$G_DCSYNC2  = '1131f6ad-9c07-11d1-f79f-00c04fc2dcd2'
$G_MEMBER   = 'bf9679c0-0de6-11d0-a285-00aa003049e2'
# Principaux « sûrs » (droits attendus)
$safe = 'Domain Admins|Admins du domaine|Enterprise Admins|Administrateurs de|Schema Admins|Administrators|Administrateurs|SYSTEM|Système|Enterprise Domain Controllers|Domain Controllers|Contrôleurs de domaine|BUILTIN|NT AUTHORITY|AUTORITE NT|Key Admins|Administrateurs clés|CREATOR OWNER|CREATEUR|Self|Cert Publishers'
$dangerRights = 'GenericAll|GenericWrite|WriteDacl|WriteOwner|WriteProperty'

# Cibles à auditer : racine, AdminSDHolder, OUs, groupes privilégiés
$targets = New-Object System.Collections.Generic.List[string]
$targets.Add($domain.DistinguishedName)
$targets.Add("CN=AdminSDHolder,CN=System,$($domain.DistinguishedName)")
Get-ADOrganizationalUnit -Filter * -ResultSetSize $null | ForEach-Object { $targets.Add($_.DistinguishedName) }
foreach ($rid in 512,519,518,520) { try { $g = Get-ADGroup -Identity "$($domain.DomainSID.Value)-$rid" -ErrorAction Stop; $targets.Add($g.DistinguishedName) } catch {} }
Write-Step ("{0} objet(s) à auditer (ACL)..." -f $targets.Count)

$rows = New-Object System.Collections.Generic.List[object]
$n = 0
foreach ($dn in ($targets | Select-Object -Unique)) {
    $n++
    Write-Progress -Activity "Analyse ACL" -Status "$n / $($targets.Count)" -PercentComplete ([int](($n/[math]::Max($targets.Count,1))*100))
    $acl = $null
    try { $acl = Get-Acl -Path ("AD:$dn") -ErrorAction Stop } catch { continue }
    foreach ($ace in $acl.Access) {
        if ($ace.AccessControlType -ne 'Allow') { continue }
        $who = "$($ace.IdentityReference)"
        if ($who -match $safe) { continue }
        $rights = "$($ace.ActiveDirectoryRights)"
        $ot = "$($ace.ObjectType)"
        $droit = $null; $risque = $null
        if ($ot -eq $G_DCSYNC1 -or $ot -eq $G_DCSYNC2) { $droit = 'DCSync (réplication annuaire)'; $risque = 'CRITIQUE' }
        elseif ($ot -eq $G_RESETPWD -and $rights -match 'ExtendedRight') { $droit = 'Réinitialiser mot de passe'; $risque = 'ELEVE' }
        elseif ($rights -match 'GenericAll') { $droit = 'GenericAll (contrôle total)'; $risque = 'ELEVE' }
        elseif ($rights -match 'WriteDacl') { $droit = 'WriteDacl (modifier ACL)'; $risque = 'ELEVE' }
        elseif ($rights -match 'WriteOwner') { $droit = 'WriteOwner (prendre possession)'; $risque = 'ELEVE' }
        elseif ($ot -eq $G_MEMBER -and $rights -match 'WriteProperty|GenericWrite') { $droit = 'Écrire membres (ajout au groupe)'; $risque = 'ELEVE' }
        elseif ($rights -match 'GenericWrite') { $droit = 'GenericWrite'; $risque = 'À surveiller' }
        else { continue }
        $rows.Add([PSCustomObject]@{
            'Objet'       = $dn
            'Principal'   = $who
            'Droit'       = $droit
            'Héritage'    = "$($ace.InheritanceType)"
            'Risque'      = $risque
        })
    }
}
Write-Progress -Activity "Analyse ACL" -Completed
$report = $rows | Sort-Object @{E={ switch($_.Risque){'CRITIQUE'{0}'ELEVE'{1}default{2}} }}, 'Objet'
$total  = @($report).Count
$dcsync = @($report | Where-Object { $_.Droit -like 'DCSync*' }).Count
$reset  = @($report | Where-Object { $_.Droit -like 'Réinitialiser*' }).Count
$crit   = @($report | Where-Object { $_.Risque -eq 'CRITIQUE' }).Count
$princ  = @($report.Principal | Select-Object -Unique).Count
Write-Step ("{0} droit(s) sensible(s) | {1} DCSync | {2} reset MDP | {3} principal(aux)." -f $total,$dcsync,$reset,$princ)

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$data = if ($total -gt 0) { $report } else { ,([PSCustomObject]@{ Information = 'Aucune délégation sensible détectée' }) }
$excel = $data | Export-Excel -Path $OutputPath -WorksheetName 'Délégations' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName 'Delegations' -TableStyle 'Medium2' -Title ("Délégations & ACL sensibles - {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm')) -TitleBold -TitleSize 14 -PassThru
if ($total -gt 0) {
    $ws = $excel.Workbook.Worksheets['Délégations']; $hr = $ws.Dimension.Start.Row
    $rg = "A{0}:{1}{2}" -f ($hr+1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column), $ws.Dimension.End.Row
    # E = Risque
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$E{0}="CRITIQUE"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::MistyRose) -Bold
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$E{0}="ELEVE"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::Moccasin)
}
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Délégations & ACL sensibles'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$kpis = @(
    @('Droits sensibles détectés', $total, 'orange'),
    @('DCSync (critique)', $dcsync, 'red'),
    @('Réinitialisation MDP déléguée', $reset, 'orange'),
    @('Risque CRITIQUE', $crit, 'red'),
    @('Principaux non-admin concernés', $princ, $null)
)
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    if ($k[2] -eq 'red' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,199,206)); $wsK.Cells["C$r"].Style.Font.Bold=$true }
    elseif ($k[2] -eq 'orange' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,235,156)) } }
$wsK.Column(2).Width=42; $wsK.Column(3).Width=16
Close-ExcelPackage $excel
Write-Host ("`n[OK] Rapport généré : {0}" -f $OutputPath) -ForegroundColor Green
