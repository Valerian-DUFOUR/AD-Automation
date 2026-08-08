<#
.SYNOPSIS
    Rapport AD CS (PKI) : autorités de certification et modèles de certificats,
    avec indicateurs de mauvaise configuration (ESC), et page KPI. Export Excel.

.DESCRIPTION
    Lit la partition de configuration de l'AD pour inventorier les autorités de
    certification (enrollmentService) et les modèles de certificats
    (pKICertificateTemplate). Signale des INDICATEURS de risque courants :
        - ENROLLEE_SUPPLIES_SUBJECT + EKU d'authentification client + pas
          d'approbation manager  -> ESC1 possible (usurpation via SAN).
        - EKU "Any Purpose" / absence d'EKU.
    Dates au format JJ-MM-AAAA.

    ATTENTION : ce sont des INDICATEURS basés sur la configuration des modèles.
    Une exploitation réelle dépend aussi des droits d'inscription (ACL). À
    confronter avec un outil dédié (Certify/Certipy) pour confirmation.

.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-ADCSReport.ps1

.NOTES
    Prérequis : RSAT ActiveDirectory + ImportExcel (auto-installé).
    Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_ADCS_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
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

$conf = (Get-ADRootDSE).configurationNamingContext
$pkiBase = "CN=Public Key Services,CN=Services,$conf"

# Autorités de certification
Write-Step "Lecture des autorités de certification..."
$cas = @()
try {
    $cas = Get-ADObject -SearchBase "CN=Enrollment Services,$pkiBase" -LDAPFilter '(objectClass=pKIEnrollmentService)' -Properties dNSHostName,cn,whenCreated -ErrorAction Stop |
        ForEach-Object { [PSCustomObject]@{ 'Autorité'=$_.cn; 'Serveur'=$_.dNSHostName; 'Date de création'=$_.whenCreated } }
} catch { Write-Warning "Pas d'AD CS détecté ou accès refusé : $($_.Exception.Message)" }

# Modèles de certificats
Write-Step "Lecture des modèles de certificats..."
$EKU_CLIENT = @('1.3.6.1.5.5.7.3.2','1.3.6.1.5.2.3.4','1.3.6.1.4.1.311.20.2.2','2.5.29.37.0')  # ClientAuth, PKINIT, SmartCard, AnyPurpose
$templates = @()
try {
    $templates = Get-ADObject -SearchBase "CN=Certificate Templates,$pkiBase" -LDAPFilter '(objectClass=pKICertificateTemplate)' `
        -Properties displayName,'msPKI-Certificate-Name-Flag','msPKI-Enrollment-Flag','msPKI-RA-Signature',pKIExtendedKeyUsage,whenCreated -ErrorAction Stop
} catch { Write-Warning "Modèles inaccessibles : $($_.Exception.Message)" }

$tplReport = foreach ($t in $templates) {
    $nameFlag = [int64]($t.'msPKI-Certificate-Name-Flag'); $enrollFlag = [int64]($t.'msPKI-Enrollment-Flag')
    $ra = [int]($t.'msPKI-RA-Signature')
    $suppliesSubject = (($nameFlag -band 0x1) -ne 0)       # ENROLLEE_SUPPLIES_SUBJECT
    $mgrApproval     = (($enrollFlag -band 0x2) -ne 0)     # PEND_ALL_REQUESTS
    $ekus = @($t.pKIExtendedKeyUsage)
    $clientAuth = ($ekus.Count -eq 0) -or ([bool]($ekus | Where-Object { $EKU_CLIENT -contains $_ }))
    $anyPurpose = ($ekus -contains '2.5.29.37.0') -or ($ekus.Count -eq 0)
    $esc1 = $suppliesSubject -and $clientAuth -and -not $mgrApproval -and ($ra -le 0)
    $risque = if ($esc1) { 'ESC1 possible' } elseif ($anyPurpose) { 'EKU Any Purpose' } else { 'OK' }
    [PSCustomObject]@{
        'Modèle'                 = if ($t.displayName) { $t.displayName } else { $t.Name }
        'Sujet fourni par demandeur' = if ($suppliesSubject) { 'Oui' } else { 'Non' }
        'Auth client (EKU)'      = if ($clientAuth) { 'Oui' } else { 'Non' }
        'Approbation manager'    = if ($mgrApproval) { 'Oui' } else { 'Non' }
        'Signatures RA requises' = $ra
        'Risque'                 = $risque
        'Date de création'       = $t.whenCreated
    }
}
$tplReport = @($tplReport) | Sort-Object @{E={ switch($_.Risque){'ESC1 possible'{0}'EKU Any Purpose'{1}default{2}} }}, 'Modèle'

$nbTpl=@($tplReport).Count; $nbEsc1=@($tplReport | Where-Object { $_.Risque -eq 'ESC1 possible' }).Count; $nbAny=@($tplReport | Where-Object { $_.Risque -eq 'EKU Any Purpose' }).Count; $nbCa=@($cas).Count
Write-Step ("{0} AC | {1} modèle(s) | {2} ESC1 possible(s)." -f $nbCa,$nbTpl,$nbEsc1)

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$script:excel = $null
function Add-Detail { param($Data,$Sheet,$Table,$DateCols)
    if (@($Data).Count -eq 0) { $Data = ,([PSCustomObject]@{ Information='Aucun élément (AD CS non déployé ?)' }) }
    $script:excel = if ($null -eq $script:excel) { $Data | Export-Excel -Path $OutputPath -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
                    else { $Data | Export-Excel -ExcelPackage $script:excel -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
    if ($DateCols) { $ws=$script:excel.Workbook.Worksheets[$Sheet]; $hr=$ws.Dimension.Start.Row; for($c=1;$c -le $ws.Dimension.End.Column;$c++){ if($DateCols -contains $ws.Cells[$hr,$c].Value){ $l=[OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c); $ws.Cells[("{0}{1}:{0}{2}" -f $l,($hr+1),$ws.Dimension.End.Row)].Style.Numberformat.Format='dd-mm-yyyy hh:mm' } } }
}
Add-Detail -Data $tplReport -Sheet 'Modèles' -Table 'Modeles' -DateCols @('Date de création')
if (@($tplReport).Count -gt 0) {
    $ws=$script:excel.Workbook.Worksheets['Modèles']; $hr=$ws.Dimension.Start.Row
    $rg = "A{0}:{1}{2}" -f ($hr+1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column), $ws.Dimension.End.Row
    # F = Risque
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$F{0}="ESC1 possible"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::MistyRose) -Bold
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$F{0}="EKU Any Purpose"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::Moccasin)
}
Add-Detail -Data $cas -Sheet 'Autorités' -Table 'AutoritesCA' -DateCols @('Date de création')
$excel = $script:excel

$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'AD CS - Autorités & modèles de certificats'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$wsK.Cells['B5'].Value = "Indicateurs de config (à confirmer avec Certify/Certipy)."
$kpis = @(
    @('Autorités de certification', $nbCa, $null),
    @('Modèles de certificats', $nbTpl, $null),
    @('ESC1 possible', $nbEsc1, 'red'),
    @('EKU Any Purpose', $nbAny, 'orange')
)
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    if ($k[2] -eq 'red' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,199,206)); $wsK.Cells["C$r"].Style.Font.Bold=$true }
    elseif ($k[2] -eq 'orange' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,235,156)) } }
$wsK.Column(2).Width=40; $wsK.Column(3).Width=16
Close-ExcelPackage $excel
Write-Host ("`n[OK] Rapport généré : {0}" -f $OutputPath) -ForegroundColor Green
