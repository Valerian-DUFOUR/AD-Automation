<#
.SYNOPSIS
    Scanner SNMP v1 / v2c : detecte les equipements repondant en SNMP et les
    communautes acceptees (dont les communautes faibles type "public"/"private"),
    avec une page KPI. Export Excel (.xlsx) sur le Bureau.

.DESCRIPTION
    Pour chaque hote cible, le script envoie une requete SNMP GET (sysDescr,
    sysName) en version 1 puis 2c, pour chaque communaute testee. Un hote qui
    repond expose potentiellement des informations sensibles ; une communaute
    "public"/"private" ou l'usage de SNMPv1 constituent des risques (SNMP v1/v2c
    n'offre aucun chiffrement ni authentification forte).

    Encodage SNMP realise en PowerShell natif (aucun module externe requis).

.PARAMETER Subnets
    Prefixes /24 a balayer (ex : -Subnets "192.168.1","10.0.0"). Scanne .1 a .254.

.PARAMETER Targets
    Liste d'IP explicites a tester (en plus / a la place des -Subnets).

.PARAMETER Communities
    Communautes SNMP a tester (defaut : public, private).

.PARAMETER TimeoutMs
    Delai d'attente par requete (defaut : 300 ms).

.PARAMETER OutputPath
    Chemin du .xlsx. Par defaut : le Bureau.

.EXAMPLE
    .\Invoke-SNMPScanner.ps1 -Subnets "192.168.1"
    .\Invoke-SNMPScanner.ps1 -Targets 192.168.1.10,192.168.1.20 -Communities public,private,community

.NOTES
    Prerequis : module ImportExcel (auto-installe). A n'utiliser que sur VOTRE
    reseau, avec autorisation. Auteur : Valerian DUFOUR / Claude
#>

[CmdletBinding()]
param(
    [string[]]$Subnets,
    [string[]]$Targets,
    [string[]]$Communities = @('public','private'),
    [int]$TimeoutMs = 300,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_SNMP_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
)

$ErrorActionPreference = 'Stop'
$Author = 'Valerian DUFOUR / Claude'
function Write-Step { param([string]$m) Write-Host "[+] $m" -ForegroundColor Cyan }

if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Step "Installation du module ImportExcel..."
    try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop }
    catch { throw "Installez ImportExcel : Install-Module ImportExcel -Scope CurrentUser" }
}
Import-Module ImportExcel -ErrorAction Stop

# --------------------------------------------------------------------------
# Encodage BER / SNMP
# --------------------------------------------------------------------------
function Encode-Len([int]$n) {
    if ($n -lt 128) { return ,([byte]$n) }
    $b = @(); $v = $n
    while ($v -gt 0) { $b = ,([byte]($v -band 0xFF)) + $b; $v = $v -shr 8 }
    return ,([byte](0x80 -bor $b.Count)) + $b
}
function Encode-TLV([byte]$tag, [byte[]]$val) { return ,([byte]$tag) + (Encode-Len $val.Count) + $val }
function Encode-Int([int]$n) {
    $bytes = [System.BitConverter]::GetBytes([int]$n); [array]::Reverse($bytes)
    $i = 0
    while ($i -lt 3 -and $bytes[$i] -eq 0 -and (($bytes[$i+1] -band 0x80) -eq 0)) { $i++ }
    return (Encode-TLV 0x02 ([byte[]]($bytes[$i..3])))
}
function Encode-OID([string]$oid) {
    $p = $oid.Split('.') | ForEach-Object { [int]$_ }
    $b = New-Object System.Collections.Generic.List[byte]
    $b.Add([byte](40 * $p[0] + $p[1]))
    for ($i = 2; $i -lt $p.Count; $i++) {
        $val = $p[$i]; $stack = @()
        if ($val -eq 0) { $stack = @(0) } else { while ($val -gt 0) { $stack = ,([byte]($val -band 0x7F)) + $stack; $val = $val -shr 7 } }
        for ($j = 0; $j -lt $stack.Count; $j++) {
            if ($j -lt $stack.Count - 1) { $b.Add([byte]($stack[$j] -bor 0x80)) } else { $b.Add([byte]$stack[$j]) }
        }
    }
    return (Encode-TLV 0x06 $b.ToArray())
}
function New-SnmpGet([int]$version, [string]$community, [string]$oid, [int]$reqId) {
    $ver  = Encode-Int $version
    $comm = Encode-TLV 0x04 ([Text.Encoding]::ASCII.GetBytes($community))
    $vb   = Encode-TLV 0x30 ((Encode-OID $oid) + (Encode-TLV 0x05 @()))
    $vbl  = Encode-TLV 0x30 $vb
    $pdu  = Encode-TLV 0xA0 ((Encode-Int $reqId) + (Encode-Int 0) + (Encode-Int 0) + $vbl)
    return (Encode-TLV 0x30 ($ver + $comm + $pdu))
}
function Read-TLV([byte[]]$b, [int]$pos) {
    $tag = $b[$pos]; $pos++
    $len = $b[$pos]; $pos++
    if ($len -band 0x80) { $n = $len -band 0x7F; $len = 0; for ($k = 0; $k -lt $n; $k++) { $len = ($len -shl 8) -bor $b[$pos]; $pos++ } }
    $val = if ($len -gt 0) { $b[$pos..($pos+$len-1)] } else { @() }
    return [PSCustomObject]@{ Tag = $tag; Val = $val; Next = ($pos + $len); ValStart = $pos }
}
function Parse-SnmpValue([byte[]]$resp) {
    try {
        $seq = Read-TLV $resp 0;        $p = $seq.ValStart
        $ver = Read-TLV $resp $p;       $p = $ver.Next
        $com = Read-TLV $resp $p;       $p = $com.Next
        $pdu = Read-TLV $resp $p;       $q = $pdu.ValStart
        $rid = Read-TLV $resp $q;       $q = $rid.Next
        $es  = Read-TLV $resp $q;       $q = $es.Next
        $ei  = Read-TLV $resp $q;       $q = $ei.Next
        $vbl = Read-TLV $resp $q;       $r = $vbl.ValStart
        $vb  = Read-TLV $resp $r;       $s = $vb.ValStart
        $oid = Read-TLV $resp $s;       $s = $oid.Next
        $val = Read-TLV $resp $s
        switch ($val.Tag) {
            0x04 { return ([Text.Encoding]::UTF8.GetString([byte[]]$val.Val)).Trim() }
            0x02 { $n = 0; foreach ($bb in $val.Val) { $n = ($n -shl 8) -bor $bb }; return "$n" }
            default { return '' }
        }
    } catch { return '' }
}
function Invoke-SnmpGet {
    param([string]$Ip, [int]$Version, [string]$Community, [string]$Oid, [int]$Timeout)
    $udp = $null
    try {
        $udp = New-Object System.Net.Sockets.UdpClient
        $udp.Client.ReceiveTimeout = $Timeout
        $pkt = New-SnmpGet $Version $Community $Oid (Get-Random -Minimum 1 -Maximum 65000)
        [void]$udp.Send($pkt, $pkt.Length, $Ip, 161)
        $ep = New-Object System.Net.IPEndPoint([System.Net.IPAddress]::Any, 0)
        $resp = $udp.Receive([ref]$ep)
        $udp.Close()
        if ($resp -and $resp.Length -gt 2 -and $resp[0] -eq 0x30) {
            return [PSCustomObject]@{ Ok = $true; Value = (Parse-SnmpValue $resp) }
        }
        return [PSCustomObject]@{ Ok = $false; Value = '' }
    } catch { if ($udp) { try { $udp.Close() } catch {} }; return [PSCustomObject]@{ Ok = $false; Value = '' } }
}

# --------------------------------------------------------------------------
# Construction de la liste de cibles
# --------------------------------------------------------------------------
$ipList = New-Object System.Collections.Generic.List[string]
foreach ($t in @($Targets)) { if ($t) { $ipList.Add($t) } }
foreach ($s in @($Subnets)) { if ($s) { for ($h = 1; $h -le 254; $h++) { $ipList.Add("$s.$h") } } }
$ipList = $ipList | Select-Object -Unique
if (@($ipList).Count -eq 0) { throw "Aucune cible. Utilisez -Subnets ou -Targets." }
Write-Step ("{0} cible(s) a tester (SNMP v1/v2c, communautes: {1})." -f @($ipList).Count, ($Communities -join ', '))

$OID_DESCR = '1.3.6.1.2.1.1.1.0'
$OID_NAME  = '1.3.6.1.2.1.1.5.0'

# --------------------------------------------------------------------------
# Scan
# --------------------------------------------------------------------------
$results = New-Object System.Collections.Generic.List[object]
$done = 0; $total = @($ipList).Count
foreach ($ip in $ipList) {
    $done++
    Write-Progress -Activity "Scan SNMP" -Status "$ip ($done/$total)" -PercentComplete ([int](($done/$total)*100))
    $hitVersion = $null; $hitComm = $null; $descr = ''; $name = ''; $versionsOk = @()
    foreach ($comm in $Communities) {
        foreach ($ver in @(0,1)) {   # 0 = SNMPv1, 1 = SNMPv2c
            $r = Invoke-SnmpGet -Ip $ip -Version $ver -Community $comm -Oid $OID_DESCR -Timeout $TimeoutMs
            if ($r.Ok) {
                $verLabel = if ($ver -eq 0) { 'v1' } else { 'v2c' }
                if ($versionsOk -notcontains $verLabel) { $versionsOk += $verLabel }
                if (-not $hitComm) {
                    $hitComm = $comm; $hitVersion = $verLabel; $descr = $r.Value
                    $rn = Invoke-SnmpGet -Ip $ip -Version $ver -Community $comm -Oid $OID_NAME -Timeout $TimeoutMs
                    if ($rn.Ok) { $name = $rn.Value }
                }
            }
        }
    }
    if ($hitComm) {
        $weak = ($hitComm -in @('public','private'))
        $risque = if ($weak -and ($versionsOk -contains 'v1')) { 'CRITIQUE' }
                  elseif ($weak -or ($versionsOk -contains 'v1')) { 'ELEVE' } else { 'MOYEN' }
        $results.Add([PSCustomObject]@{
            'IP'               = $ip
            'Versions'         = ($versionsOk -join ' / ')
            'Communaute'       = $hitComm
            'Communaute faible'= if ($weak) { 'Oui' } else { 'Non' }
            'sysName'          = $name
            'sysDescr'         = $descr
            'Risque'           = $risque
        })
        Write-Host ("   [SNMP] {0,-16} {1,-6} comm='{2}' {3}" -f $ip, ($versionsOk -join '/'), $hitComm, $risque) -ForegroundColor Yellow
    }
}
Write-Progress -Activity "Scan SNMP" -Completed
$report = $results | Sort-Object 'Risque','IP'
Write-Step ("{0} hote(s) SNMP repondant(s) sur {1} teste(s)." -f @($report).Count, $total)

# --------------------------------------------------------------------------
# KPI
# --------------------------------------------------------------------------
$nb      = @($report).Count
$nbV1    = @($report | Where-Object { $_.Versions -match 'v1' }).Count
$nbV2    = @($report | Where-Object { $_.Versions -match 'v2c' }).Count
$nbWeak  = @($report | Where-Object { $_.'Communaute faible' -eq 'Oui' }).Count
$nbCrit  = @($report | Where-Object { $_.Risque -eq 'CRITIQUE' }).Count

# --------------------------------------------------------------------------
# Export Excel
# --------------------------------------------------------------------------
Write-Step "Generation du fichier Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$data = if ($nb -gt 0) { $report } else { ,([PSCustomObject]@{ Information = 'Aucun hote SNMP detecte' }) }
$excel = $data | Export-Excel -Path $OutputPath -WorksheetName 'SNMP' -AutoSize -AutoFilter `
    -FreezeTopRow -BoldTopRow -TableName 'SNMP' -TableStyle 'Medium2' `
    -Title ("Scan SNMP v1/v2c - {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm')) -TitleBold -TitleSize 14 -PassThru

if ($nb -gt 0) {
    $ws = $excel.Workbook.Worksheets['SNMP']; $hr = $ws.Dimension.Start.Row
    $rg = "A{0}:{1}{2}" -f ($hr+1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column), $ws.Dimension.End.Row
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$G{0}="CRITIQUE"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::MistyRose) -Bold
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$G{0}="ELEVE"' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::Moccasin)
}

$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Scan SNMP v1 / v2c'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$wsK.Cells['B5'].Value = ("Cibles testees : {0}" -f $total)
$kpis = @(
    @('Hotes SNMP repondants', $nb, $null),
    @('  dont SNMPv1 (obsolete)', $nbV1, 'red'),
    @('  dont SNMPv2c', $nbV2, $null),
    @('Communautes faibles (public/private)', $nbWeak, 'red'),
    @('Risque CRITIQUE', $nbCrit, 'red')
)
$r = 7
$wsK.Cells["B$r"].Value = 'Indicateur'; $wsK.Cells["C$r"].Value = 'Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold = $true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) {
    $r++; $wsK.Cells["B$r"].Value = $k[0]; $wsK.Cells["C$r"].Value = $k[1]
    if ($k[2] -eq 'red' -and $k[1] -gt 0) {
        $wsK.Cells["C$r"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,199,206))
        $wsK.Cells["C$r"].Style.Font.Bold = $true
    }
}
$wsK.Column(2).Width = 42; $wsK.Column(3).Width = 16
Close-ExcelPackage $excel

Write-Host ""
Write-Host ("[OK] Rapport genere : {0}" -f $OutputPath) -ForegroundColor Green
Write-Host ("     {0} hote(s) SNMP | {1} communaute(s) faible(s) | {2} critique(s)." -f $nb, $nbWeak, $nbCrit) -ForegroundColor Green
