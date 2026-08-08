<#
.SYNOPSIS
    Rapport des connexions LDAP NON sécurisées (binds simples/en clair ou non
    signés) vers les contrôleurs de domaine, avec page KPI. Export Excel (.xlsx).

.DESCRIPTION
    Analyse le journal « Directory Service » des contrôleurs de domaine pour
    identifier les clients et comptes utilisant encore du LDAP non sécurisé,
    avant d'imposer la signature LDAP / le channel binding :
        - Event 2889 : détail par client (IP + identité + type de bind).
        - Event 2887 : synthèse périodique du nombre de binds non sécurisés.

    Regroupe par compte et par adresse IP. Dates au format JJ-MM-AAAA.

    IMPORTANT : l'event 2889 n'est journalisé que si la journalisation
    diagnostique « 16 LDAP Interface Events » est réglée sur 2 sur les DC :
        reg add "HKLM\SYSTEM\CurrentControlSet\Services\NTDS\Diagnostics" ^
            /v "16 LDAP Interface Events" /t REG_DWORD /d 2 /f
    Sans cela, seul l'event 2887 (compteur global) est disponible.

.PARAMETER Days
    Fenêtre d'analyse HISTORIQUE en jours (défaut : 14). Utilisé si -MonitorMinutes = 0.
.PARAMETER MonitorMinutes
    Mode CAPTURE LIVE : surveille les binds non sécurisés en continu pendant N
    minutes (0 = désactivé, mode historique). Idéal avant d'imposer la signature
    LDAP, pour être sûr de ne rien rater.
.PARAMETER IntervalSeconds
    Intervalle de relève en mode capture live (défaut : 30 s).
.PARAMETER DomainControllers
    Liste de DC à interroger (défaut : tous les DC du domaine).
.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-InsecureLDAPReport.ps1 -Days 30

.NOTES
    Prérequis : RSAT ActiveDirectory + ImportExcel (auto-installé) ; droits de
    lecture des journaux d'événements des DC. À lancer depuis un DC ou un poste
    d'administration. Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [int]$Days = 14,
    [int]$MonitorMinutes = 0,
    [int]$IntervalSeconds = 30,
    [string[]]$DomainControllers,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_LDAP_NonSecurise_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
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

if (-not $DomainControllers) {
    try { $DomainControllers = (Get-ADDomainController -Filter *).HostName } catch { $DomainControllers = @($env:COMPUTERNAME) }
}
$events = New-Object System.Collections.Generic.List[object]
$summaryTotal = 0
$seen = @{}
function Add-Event2889 { param($e,$dc)
    $key = "$dc-$($e.RecordId)"; if ($seen.ContainsKey($key)) { return }; $seen[$key] = $true
    $ip = ("$($e.Properties[0].Value)" -split ':')[0]
    $ident = "$($e.Properties[1].Value)"
    $typeLbl = if ("$($e.Properties[2].Value)" -eq '1') { 'Simple (mot de passe en clair)' } else { 'Non signé (SASL)' }
    $events.Add([PSCustomObject]@{ DC=$dc; IP=$ip; Compte=$ident; Type=$typeLbl; Date=$e.TimeCreated })
}

if ($MonitorMinutes -gt 0) {
    $end = (Get-Date).AddMinutes($MonitorMinutes)
    Write-Step ("Capture LIVE pendant {0} min (relève toutes les {1}s) sur {2} DC..." -f $MonitorMinutes, $IntervalSeconds, @($DomainControllers).Count)
    while ((Get-Date) -lt $end) {
        foreach ($dc in $DomainControllers) {
            try { Get-WinEvent -ComputerName $dc -FilterHashtable @{ LogName='Directory Service'; Id=2889; StartTime=(Get-Date).AddMinutes(-6) } -ErrorAction Stop | ForEach-Object { Add-Event2889 $_ $dc } } catch {}
        }
        $remain = [int]($end - (Get-Date)).TotalSeconds
        Write-Progress -Activity "Capture LDAP live" -Status ("{0} bind(s) capturé(s) - {1}s restantes" -f $events.Count, [math]::Max($remain,0)) -PercentComplete ([int]((($MonitorMinutes*60 - [math]::Max($remain,0))/[double]($MonitorMinutes*60))*100))
        if ((Get-Date) -lt $end) { Start-Sleep -Seconds ([math]::Min($IntervalSeconds, [math]::Max($remain,1))) }
    }
    Write-Progress -Activity "Capture LDAP live" -Completed
} else {
    $after = (Get-Date).AddDays(-$Days)
    Write-Step ("Analyse historique sur {0} DC depuis le {1}." -f @($DomainControllers).Count, $after.ToString('dd-MM-yyyy'))
    foreach ($dc in $DomainControllers) {
        try { Get-WinEvent -ComputerName $dc -FilterHashtable @{ LogName='Directory Service'; Id=2889; StartTime=$after } -ErrorAction Stop | ForEach-Object { Add-Event2889 $_ $dc } } catch { Write-Warning "2889 indisponible sur $dc ($($_.Exception.Message))" }
        try { $s = Get-WinEvent -ComputerName $dc -FilterHashtable @{ LogName='Directory Service'; Id=2887; StartTime=$after } -ErrorAction Stop
            foreach ($e in $s) { foreach ($p in $e.Properties) { $n = 0; if ([int]::TryParse("$($p.Value)",[ref]$n)) { $summaryTotal += $n } } } } catch {}
    }
}

# Agrégations
$byClient = $events | Group-Object IP,Compte,Type | ForEach-Object {
    $f = $_.Group[0]
    [PSCustomObject]@{
        'Adresse IP'          = $f.IP
        'Compte'              = $f.Compte
        'Type de bind'        = $f.Type
        'Occurrences'         = $_.Count
        'Dernière occurrence' = ($_.Group | Sort-Object Date -Descending | Select-Object -First 1).Date
        'DC'                  = ($_.Group.DC | Select-Object -Unique) -join ', '
    }
} | Sort-Object Occurrences -Descending

$byAccount = $events | Group-Object Compte | ForEach-Object {
    [PSCustomObject]@{ 'Compte'=$_.Name; 'Occurrences'=$_.Count; 'IP distinctes'=@($_.Group.IP | Select-Object -Unique).Count }
} | Sort-Object Occurrences -Descending

$nbEvents  = @($events).Count
$nbCompte  = @($byAccount).Count
$nbIP      = @($events.IP | Select-Object -Unique).Count
$nbClair   = @($events | Where-Object { $_.Type -like 'Simple*' }).Count

Write-Step ("{0} bind(s) non sécurisé(s) détaillé(s) | {1} compte(s) | {2} IP." -f $nbEvents,$nbCompte,$nbIP)

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$script:excel = $null
function Add-Detail { param($Data,$Sheet,$Table,$DateCols)
    if (@($Data).Count -eq 0) { $Data = ,([PSCustomObject]@{ Information='Aucun élément (journalisation 2889 activée ?)' }) }
    $script:excel = if ($null -eq $script:excel) { $Data | Export-Excel -Path $OutputPath -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
                    else { $Data | Export-Excel -ExcelPackage $script:excel -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
    if ($DateCols) { $ws=$script:excel.Workbook.Worksheets[$Sheet]; $hr=$ws.Dimension.Start.Row; for($c=1;$c -le $ws.Dimension.End.Column;$c++){ if($DateCols -contains $ws.Cells[$hr,$c].Value){ $l=[OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c); $ws.Cells[("{0}{1}:{0}{2}" -f $l,($hr+1),$ws.Dimension.End.Row)].Style.Numberformat.Format='dd-mm-yyyy hh:mm' } } }
}
Add-Detail -Data $byClient  -Sheet 'Binds non sécurisés' -Table 'BindsLDAP' -DateCols @('Dernière occurrence')
Add-Detail -Data $byAccount -Sheet 'Par compte' -Table 'ParCompte'
$excel = $script:excel

$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'LDAP non sécurisé (binds simples / non signés)'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$modeLbl = if ($MonitorMinutes -gt 0) { "Capture live $MonitorMinutes min" } else { "Historique $Days j" }
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}  |  Mode : {1}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'), $modeLbl)
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
if ($nbEvents -eq 0) { $wsK.Cells['B5'].Value = "Aucun event 2889 : activez « 16 LDAP Interface Events = 2 » sur les DC (voir en-tête du script). Compteur 2887 = $summaryTotal." }
$kpis = @(
    @('Binds non sécurisés (détail 2889)', $nbEvents, 'red'),
    @('Comptes concernés', $nbCompte, 'orange'),
    @('Adresses IP concernées', $nbIP, 'orange'),
    @('  dont binds en clair (simple)', $nbClair, 'red'),
    @('Compteur global (event 2887)', $summaryTotal, $null)
)
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    if ($k[2] -eq 'red' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,199,206)); $wsK.Cells["C$r"].Style.Font.Bold=$true }
    elseif ($k[2] -eq 'orange' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,235,156)) } }
$wsK.Column(2).Width=42; $wsK.Column(3).Width=18
Close-ExcelPackage $excel
Write-Host ("`n[OK] Rapport généré : {0}" -f $OutputPath) -ForegroundColor Green
