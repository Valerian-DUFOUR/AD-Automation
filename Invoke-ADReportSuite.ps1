<#
.SYNOPSIS
    Chef d'orchestre : exécute l'ensemble des rapports AD dans un dossier daté,
    produit une synthèse d'exécution (KPI) et, en option, envoie le tout par
    e-mail. À planifier (tâche planifiée) pour un reporting récurrent.

.DESCRIPTION
    Lance, depuis le dossier du script, la série des rapports d'extraction et de
    sécurité AD (chacun génère son .xlsx avec sa propre page KPI) dans un dossier
    horodaté, puis crée « 00_Synthese_Execution.xlsx » récapitulant le statut et
    la durée de chaque rapport. Avec -SendEmail, envoie les fichiers par SMTP.

.PARAMETER OutputDir
    Dossier de sortie (défaut : Bureau\AD_Reports_<date>).
.PARAMETER Scripts
    Liste de scripts à exécuter (défaut : jeu AD complet, hors scan réseau/SNMP).
.PARAMETER SendEmail
    Envoie les rapports par e-mail (nécessite -SmtpServer, -From, -To).
.PARAMETER SmtpServer / -From / -To / -Port / -UseSsl
    Paramètres SMTP pour l'envoi.

.EXAMPLE
    .\Invoke-ADReportSuite.ps1
    .\Invoke-ADReportSuite.ps1 -SendEmail -SmtpServer smtp.example.local -From ad@example.local -To rssi@example.local

.NOTES
    Prérequis : les scripts Export-*/… présents dans le même dossier, RSAT +
    ImportExcel. L'envoi utilise Send-MailMessage (déprécié mais fonctionnel ;
    aucun mot de passe n'est stocké). Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string]$OutputDir = (Join-Path ([Environment]::GetFolderPath('Desktop')) ("AD_Reports_{0}" -f (Get-Date -Format 'dd-MM-yyyy_HHmm'))),
    [string[]]$Scripts,
    [switch]$SendEmail,
    [string]$SmtpServer,
    [string]$From,
    [string[]]$To,
    [int]$Port = 25,
    [switch]$UseSsl
)
$ErrorActionPreference = 'Continue'
$Author = 'Valérian DUFOUR / Claude'
function Write-Step { param($m) Write-Host "[+] $m" -ForegroundColor Cyan }
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop } catch { throw "Install-Module ImportExcel -Scope CurrentUser" }
}
Import-Module ImportExcel -ErrorAction Stop
if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }

if (-not $Scripts) {
    $Scripts = @(
        'Export-ADUsersReport','Export-ADComputersReport','Export-ADGroupsReport','Export-GPOReport',
        'Export-ADDuplicateObjects','Export-ADSecurityPosture','Export-ADInactiveReport','Export-ADComputersNoBitLocker',
        'Export-LAPSReport','Export-ADTrustsReport','Export-ADDelegationReport','Export-GPOSecurityAudit',
        'Export-ADCSReport','Export-ADHealthReport','Export-ADExpirationReport','Export-InsecureLDAPReport',
        'Export-DHCP-DNSReport','Export-ADChangeDiff'
    )
}
Write-Step ("Exécution de {0} rapport(s) -> {1}" -f $Scripts.Count, $OutputDir)

$results = New-Object System.Collections.Generic.List[object]
$globalStart = Get-Date
foreach ($name in $Scripts) {
    $path = Join-Path $PSScriptRoot "$name.ps1"
    $file = Join-Path $OutputDir "$name.xlsx"
    if (-not (Test-Path $path)) {
        $results.Add([PSCustomObject]@{ 'Rapport'=$name; 'Statut'='INTROUVABLE'; 'Durée (s)'=0; 'Fichier'='' }); continue
    }
    Write-Host ("--- $name ---") -ForegroundColor DarkCyan
    $t0 = Get-Date; $status = 'OK'
    try { & $path -OutputPath $file } catch { $status = "ERREUR: $($_.Exception.Message)" }
    $dur = [math]::Round(((Get-Date) - $t0).TotalSeconds,1)
    if ($status -eq 'OK' -and -not (Test-Path $file)) { $status = 'AUCUN FICHIER' }
    $results.Add([PSCustomObject]@{ 'Rapport'=$name; 'Statut'=$status; 'Durée (s)'=$dur; 'Fichier'=(if(Test-Path $file){Split-Path $file -Leaf}else{''}) })
}
$totalDur = [math]::Round(((Get-Date) - $globalStart).TotalMinutes,1)
$ok = @($results | Where-Object { $_.Statut -eq 'OK' }).Count
$ko = @($results | Where-Object { $_.Statut -ne 'OK' }).Count

# --- Synthèse d'exécution ---
$summaryFile = Join-Path $OutputDir '00_Synthese_Execution.xlsx'
if (Test-Path $summaryFile) { Remove-Item $summaryFile -Force }
$excel = $results | Export-Excel -Path $summaryFile -WorksheetName 'Exécution' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName 'Execution' -TableStyle 'Medium2' -Title ("Synthèse d'exécution - {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm')) -TitleBold -TitleSize 14 -PassThru
$ws = $excel.Workbook.Worksheets['Exécution']; $hr = $ws.Dimension.Start.Row
$rg = "A{0}:{1}{2}" -f ($hr+1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column), $ws.Dimension.End.Row
Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$B{0}="OK"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::FromArgb(198,239,206))
Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=LEFT($B{0},6)="ERREUR"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::FromArgb(255,199,206)) -Bold
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Suite de rapports Active Directory'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date : {0}  |  Durée totale : {1} min" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'), $totalDur)
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$wsK.Cells['B5'].Value = ("Dossier : {0}" -f $OutputDir)
$kpis = @(
    @('Rapports demandés', $Scripts.Count, $null),
    @('Réussis', $ok, 'green'),
    @('En échec / absents', $ko, 'red'),
    @('Durée totale (min)', $totalDur, $null)
)
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    $col=$null; switch($k[2]){'red'{$col=[System.Drawing.Color]::FromArgb(255,199,206)}'green'{$col=[System.Drawing.Color]::FromArgb(198,239,206)}}
    if ($col -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor($col); $wsK.Cells["C$r"].Style.Font.Bold=$true } }
$wsK.Column(2).Width=36; $wsK.Column(3).Width=18
Close-ExcelPackage $excel
Write-Step ("Synthèse : {0}  ({1} OK / {2} KO, {3} min)" -f $summaryFile, $ok, $ko, $totalDur)

# --- Envoi e-mail (optionnel) ---
if ($SendEmail) {
    if (-not $SmtpServer -or -not $From -or -not $To) { Write-Warning "SendEmail nécessite -SmtpServer, -From et -To." }
    else {
        $attach = Get-ChildItem -Path $OutputDir -Filter *.xlsx | Select-Object -ExpandProperty FullName
        $body = "Rapports AD du {0}.`n{1} rapport(s) OK, {2} en échec. Durée {3} min.`nDossier : {4}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'), $ok, $ko, $totalDur, $OutputDir
        try {
            $p = @{ SmtpServer=$SmtpServer; From=$From; To=$To; Port=$Port; Subject=("[AD] Rapports du {0}" -f (Get-Date -Format 'dd-MM-yyyy')); Body=$body; Attachments=$attach }
            if ($UseSsl) { $p['UseSsl'] = $true }
            Send-MailMessage @p -ErrorAction Stop
            Write-Step "E-mail envoyé à : $($To -join ', ')"
        } catch { Write-Warning "Échec envoi e-mail : $($_.Exception.Message)" }
    }
}
Write-Host ("`n[OK] Suite terminée : {0}" -f $OutputDir) -ForegroundColor Green
