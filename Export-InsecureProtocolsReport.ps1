<#
.SYNOPSIS
    Audit des protocoles et réglages vulnérables / obsolètes d'un serveur Windows
    (SMBv1, signatures SMB, LLMNR, NTLMv1/LM, WDigest, TLS/SSL obsolètes, LDAP
    signing, RDP NLA, PowerShell v2, Telnet/FTP…), avec page KPI. Export Excel.

.DESCRIPTION
    Vérifie la configuration (registre / services) de la machine locale — ou de
    plusieurs machines via -ComputerName — face à une liste de protocoles
    dangereux ou dépréciés, et classe chaque point en OK / À risque / Critique.
    Idéal exécuté sur les contrôleurs de domaine et les serveurs.

.PARAMETER ComputerName
    Machines distantes à auditer (via WinRM). Défaut : machine locale.
.PARAMETER CaptureMinutes
    Capture LIVE de l'USAGE réel de SMBv1 pendant N minutes (0 = désactivé), afin
    de repérer les clients qui utilisent encore SMBv1 avant de le désactiver.
    Nécessite l'audit SMBv1 activé : Set-SmbServerConfiguration -AuditSmb1Access $true
.PARAMETER IntervalSeconds
    Intervalle de relève en mode capture (défaut : 30 s).
.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-InsecureProtocolsReport.ps1
    .\Export-InsecureProtocolsReport.ps1 -ComputerName DC01,SRV01,SRV02

.NOTES
    Prérequis : ImportExcel (auto-installé). Droits admin/lecture registre ;
    WinRM pour les machines distantes. Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string[]]$ComputerName,
    [int]$CaptureMinutes = 0,
    [int]$IntervalSeconds = 30,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_Protocoles_Vulnerables_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
)
$ErrorActionPreference = 'Stop'
$Author = 'Valérian DUFOUR / Claude'
function Write-Step { param($m) Write-Host "[+] $m" -ForegroundColor Cyan }
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Step "Installation d'ImportExcel..."; try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop } catch { throw "Install-Module ImportExcel -Scope CurrentUser" }
}
Import-Module ImportExcel -ErrorAction Stop

# Bloc de vérifications exécuté localement ou à distance
$checks = {
    function RegVal { param($Path,$Name) try { (Get-ItemProperty -Path $Path -Name $Name -ErrorAction Stop).$Name } catch { $null } }
    $out = New-Object System.Collections.Generic.List[object]
    function Add-Check { param($Reglage,$Etat,$Attendu,$Niveau,$Reco) $out.Add([PSCustomObject]@{
        Machine=$env:COMPUTERNAME; Reglage=$Reglage; 'Etat actuel'=$Etat; Attendu=$Attendu; Niveau=$Niveau; Recommandation=$Reco }) }

    # SMBv1 serveur
    $smb1 = RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters' 'SMB1'
    $smb1on = ($smb1 -ne 0)
    Add-Check 'SMBv1 (serveur)' (if ($smb1on) {'Activé'} else {'Désactivé'}) 'Désactivé' (if ($smb1on){'CRITIQUE'}else{'OK'}) 'Désactiver SMBv1 (Disable-WindowsOptionalFeature SMB1Protocol)'
    # Signature SMB requise (serveur)
    $smbsign = RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters' 'RequireSecuritySignature'
    Add-Check 'Signature SMB requise (serveur)' (if ($smbsign -eq 1){'Oui'}else{'Non'}) 'Oui' (if ($smbsign -eq 1){'OK'}else{'À risque'}) 'Exiger la signature SMB'
    # SMBv1 client (pilote)
    $mrx = RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\mrxsmb10' 'Start'
    Add-Check 'SMBv1 (client mrxsmb10)' (if ($mrx -eq 4){'Désactivé'}elseif($null -eq $mrx){'Absent'}else{'Activé'}) 'Désactivé/Absent' (if ($mrx -eq 4 -or $null -eq $mrx){'OK'}else{'À risque'}) 'Désactiver le client SMBv1'
    # LLMNR
    $llmnr = RegVal 'HKLM:\SOFTWARE\Policies\Microsoft\Windows NT\DNSClient' 'EnableMulticast'
    Add-Check 'LLMNR' (if ($llmnr -eq 0){'Désactivé'}else{'Activé/Non défini'}) 'Désactivé (0)' (if ($llmnr -eq 0){'OK'}else{'À risque'}) 'Désactiver LLMNR par GPO'
    # NTLMv1 / LM
    $lm = RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'LmCompatibilityLevel'
    Add-Check 'Niveau NTLM (LmCompatibilityLevel)' (if($null -eq $lm){'Non défini'}else{"$lm"}) '5 (NTLMv2 uniquement)' (if ($lm -ge 5){'OK'}elseif($null -eq $lm){'À risque'}else{'CRITIQUE'}) 'Fixer LmCompatibilityLevel=5'
    # WDigest
    $wd = RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\WDigest' 'UseLogonCredential'
    Add-Check 'WDigest (mots de passe en mémoire)' (if ($wd -eq 1){'Activé'}else{'Désactivé'}) 'Désactivé (0)' (if ($wd -eq 1){'CRITIQUE'}else{'OK'}) 'UseLogonCredential=0'
    # LSA Protection
    $ppl = RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Lsa' 'RunAsPPL'
    Add-Check 'Protection LSA (RunAsPPL)' (if ($ppl -eq 1){'Activé'}else{'Désactivé'}) 'Activé (1)' (if ($ppl -eq 1){'OK'}else{'À risque'}) 'Activer RunAsPPL'
    # TLS/SSL obsolètes (serveur)
    foreach ($proto in 'SSL 2.0','SSL 3.0','TLS 1.0','TLS 1.1') {
        $en = RegVal "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Server" 'Enabled'
        $dis = RegVal "HKLM:\SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\$proto\Server" 'DisabledByDefault'
        $off = ($en -eq 0) -or ($dis -eq 1)
        $niv = if ($off) {'OK'} elseif ($proto -like 'SSL*') {'CRITIQUE'} else {'À risque'}
        Add-Check "$proto (serveur)" (if($off){'Désactivé'}else{'Activé/Non défini'}) 'Désactivé' $niv "Désactiver $proto dans SCHANNEL"
    }
    # LDAP signing (DC uniquement)
    $ldapInt = RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters' 'LDAPServerIntegrity'
    if ($null -ne $ldapInt) { Add-Check 'Signature LDAP (DC)' (if($ldapInt -eq 2){'Requise'}else{"$ldapInt"}) 'Requise (2)' (if($ldapInt -eq 2){'OK'}else{'À risque'}) 'LDAPServerIntegrity=2' }
    $cbt = RegVal 'HKLM:\SYSTEM\CurrentControlSet\Services\NTDS\Parameters' 'LdapEnforceChannelBinding'
    if ($null -ne $cbt) { Add-Check 'LDAP Channel Binding (DC)' (if($cbt -ge 1){'Activé'}else{'Désactivé'}) 'Activé (2)' (if($cbt -ge 1){'OK'}else{'À risque'}) 'LdapEnforceChannelBinding=2' }
    # RDP NLA
    $nla = RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server\WinStations\RDP-Tcp' 'UserAuthentication'
    $deny = RegVal 'HKLM:\SYSTEM\CurrentControlSet\Control\Terminal Server' 'fDenyTSConnections'
    if ($deny -ne 1) { Add-Check 'RDP - authentification NLA' (if($nla -eq 1){'Activée'}else{'Désactivée'}) 'Activée' (if($nla -eq 1){'OK'}else{'À risque'}) 'Exiger NLA pour RDP' }
    # PowerShell v2
    $psv2 = $null; try { $psv2 = (Get-WindowsOptionalFeature -Online -FeatureName MicrosoftWindowsPowerShellV2 -ErrorAction Stop).State } catch {}
    if ($psv2) { Add-Check 'PowerShell v2 (moteur)' "$psv2" 'Disabled' (if("$psv2" -eq 'Disabled'){'OK'}else{'À risque'}) 'Désinstaller le moteur PowerShell v2' }
    # Services obsolètes
    foreach ($svc in @(@{N='Telnet';S='TlntSvr'},@{N='FTP (IIS)';S='FTPSVC'},@{N='Spouleur d''impression';S='Spooler'})) {
        try { $s = Get-Service -Name $svc.S -ErrorAction Stop; $niv = if ($svc.S -eq 'Spooler'){'À risque'}else{'CRITIQUE'}; Add-Check "Service $($svc.N)" (if($s.Status -eq 'Running'){'En cours'}else{"$($s.Status)"}) 'Arrêté/Désactivé (si non requis)' (if($s.Status -eq 'Running'){$niv}else{'OK'}) 'Désactiver si non nécessaire' } catch {}
    }
    return $out
}

$targets = if ($ComputerName) { $ComputerName } else { @($env:COMPUTERNAME) }
Write-Step ("Audit de {0} machine(s)..." -f @($targets).Count)
$all = New-Object System.Collections.Generic.List[object]
if ($ComputerName) {
    foreach ($m in $ComputerName) {
        try { $res = Invoke-Command -ComputerName $m -ScriptBlock $checks -ErrorAction Stop; $res | ForEach-Object { $all.Add($_) } }
        catch { Write-Warning "Échec sur $m : $($_.Exception.Message)" }
    }
} else {
    (& $checks) | ForEach-Object { $all.Add($_) }
}
$report = $all | Select-Object Machine,Reglage,'Etat actuel',Attendu,Niveau,Recommandation |
    Sort-Object @{E={ switch($_.Niveau){'CRITIQUE'{0}'À risque'{1}default{2}} }}, Machine, Reglage

$total = @($report).Count
$crit  = @($report | Where-Object { $_.Niveau -eq 'CRITIQUE' }).Count
$risk  = @($report | Where-Object { $_.Niveau -eq 'À risque' }).Count
$ok    = @($report | Where-Object { $_.Niveau -eq 'OK' }).Count
Write-Step ("{0} contrôle(s) | {1} critique(s) | {2} à risque | {3} OK." -f $total,$crit,$risk,$ok)

# --- Capture LIVE de l'usage SMBv1 (optionnelle) ---
$smb1usage = @()
if ($CaptureMinutes -gt 0) {
    $end = (Get-Date).AddMinutes($CaptureMinutes); $seen = @{}; $raw = New-Object System.Collections.Generic.List[object]
    Write-Step ("Capture LIVE de l'usage SMBv1 pendant {0} min (relève {1}s)..." -f $CaptureMinutes, $IntervalSeconds)
    while ((Get-Date) -lt $end) {
        foreach ($m in $targets) {
            try {
                $fh = @{ LogName='Microsoft-Windows-SMBServer/Audit'; Id=3000; StartTime=(Get-Date).AddMinutes(-6) }
                $evs = if ($m -eq $env:COMPUTERNAME) { Get-WinEvent -FilterHashtable $fh -ErrorAction Stop } else { Get-WinEvent -ComputerName $m -FilterHashtable $fh -ErrorAction Stop }
                foreach ($e in $evs) { $k = "$m-$($e.RecordId)"; if (-not $seen.ContainsKey($k)) { $seen[$k]=$true; $cli=''; try { $cli = "$($e.Properties[0].Value)" } catch {}; $raw.Add([PSCustomObject]@{ Machine=$m; Client=$cli; Date=$e.TimeCreated }) } }
            } catch {}
        }
        $remain = [int]($end - (Get-Date)).TotalSeconds
        Write-Progress -Activity "Capture SMBv1" -Status ("{0} accès - {1}s restantes" -f $raw.Count, [math]::Max($remain,0)) -PercentComplete ([int]((($CaptureMinutes*60 - [math]::Max($remain,0))/[double]($CaptureMinutes*60))*100))
        if ((Get-Date) -lt $end) { Start-Sleep -Seconds ([math]::Min($IntervalSeconds, [math]::Max($remain,1))) }
    }
    Write-Progress -Activity "Capture SMBv1" -Completed
    $smb1usage = $raw | Group-Object Machine,Client | ForEach-Object { $f=$_.Group[0]; [PSCustomObject]@{ 'Machine'=$f.Machine; 'Client SMBv1'=$f.Client; 'Occurrences'=$_.Count; 'Dernier accès'=($_.Group | Sort-Object Date -Descending | Select-Object -First 1).Date } } | Sort-Object Occurrences -Descending
    if (@($smb1usage).Count -eq 0) { Write-Step "Aucun accès SMBv1 capturé (audit activé ? Set-SmbServerConfiguration -AuditSmb1Access `$true)" }
    else { Write-Step ("{0} client(s) SMBv1 vu(s) durant la capture." -f @($smb1usage).Count) }
}

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$data = if ($total -gt 0) { $report } else { ,([PSCustomObject]@{ Information='Aucun contrôle' }) }
$excel = $data | Export-Excel -Path $OutputPath -WorksheetName 'Contrôles' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow `
    -TableName 'Controles' -TableStyle 'Medium2' -Title ("Protocoles vulnérables / obsolètes - {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm')) -TitleBold -TitleSize 14 -PassThru
if ($total -gt 0) {
    $ws = $excel.Workbook.Worksheets['Contrôles']; $hr = $ws.Dimension.Start.Row
    $rg = "A{0}:{1}{2}" -f ($hr+1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column), $ws.Dimension.End.Row
    # E = Niveau
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$E{0}="CRITIQUE"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::FromArgb(255,199,206)) -Bold
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$E{0}="À risque"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::FromArgb(255,235,156))
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$E{0}="OK"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::FromArgb(198,239,206))
}
if ($CaptureMinutes -gt 0) {
    $du = if (@($smb1usage).Count -gt 0) { $smb1usage } else { ,([PSCustomObject]@{ Information = 'Aucun accès SMBv1 capturé (activez l''audit : Set-SmbServerConfiguration -AuditSmb1Access $true)' }) }
    $excel = $du | Export-Excel -ExcelPackage $excel -WorksheetName 'Usage SMBv1 (capture)' -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'UsageSMB1' -TableStyle 'Medium2' -PassThru
    if (@($smb1usage).Count -gt 0) {
        $ws2 = $excel.Workbook.Worksheets['Usage SMBv1 (capture)']; $hr2 = $ws2.Dimension.Start.Row
        for ($c=1;$c -le $ws2.Dimension.End.Column;$c++){ if ($ws2.Cells[$hr2,$c].Value -eq 'Dernier accès'){ $l=[OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c); $ws2.Cells[("{0}{1}:{0}{2}" -f $l,($hr2+1),$ws2.Dimension.End.Row)].Style.Numberformat.Format='dd-mm-yyyy hh:mm' } }
    }
}
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Protocoles vulnérables / obsolètes'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}  |  Machines : {1}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'), (@($targets) -join ', '))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$kpis = @(
    @('Contrôles effectués', $total, $null),
    @('CRITIQUES', $crit, 'red'),
    @('À risque', $risk, 'orange'),
    @('Conformes (OK)', $ok, 'green')
)
if ($CaptureMinutes -gt 0) { $kpis += ,@('Clients SMBv1 vus (capture)', @($smb1usage).Count, 'red') }
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    $col=$null; switch($k[2]){'red'{$col=[System.Drawing.Color]::FromArgb(255,199,206)}'orange'{$col=[System.Drawing.Color]::FromArgb(255,235,156)}'green'{$col=[System.Drawing.Color]::FromArgb(198,239,206)}}
    if ($col -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor($col); $wsK.Cells["C$r"].Style.Font.Bold=$true } }
$gr = $r + 3
$wsK.Cells["B$gr"].Value='Niveau'; $wsK.Cells["C$gr"].Value='Nombre'; $wsK.Cells["B$gr:C$gr"].Style.Font.Bold=$true
$wsK.Cells["B$($gr+1)"].Value='Critique'; $wsK.Cells["C$($gr+1)"].Value=$crit
$wsK.Cells["B$($gr+2)"].Value='À risque'; $wsK.Cells["C$($gr+2)"].Value=$risk
$wsK.Cells["B$($gr+3)"].Value='OK'; $wsK.Cells["C$($gr+3)"].Value=$ok
try { $ch=$wsK.Drawings.AddChart('protoChart',[OfficeOpenXml.Drawing.Chart.eChartType]::Pie); $ch.Title.Text='Répartition des contrôles'; $ch.SetPosition(6,0,5,0); $ch.SetSize(400,280); $null=$ch.Series.Add($wsK.Cells["C$($gr+1):C$($gr+3)"],$wsK.Cells["B$($gr+1):B$($gr+3)"]); $ch.DataLabel.ShowValue=$true } catch {}
$wsK.Column(2).Width=42; $wsK.Column(3).Width=16
Close-ExcelPackage $excel
Write-Host ("`n[OK] Rapport généré : {0}" -f $OutputPath) -ForegroundColor Green
