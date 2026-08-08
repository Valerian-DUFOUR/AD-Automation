<#
.SYNOPSIS
    Rapport DHCP + DNS : étendues et taux d'occupation DHCP, zones DNS, avec
    page KPI. Export Excel (.xlsx) sur le Bureau.

.DESCRIPTION
    - DHCP : serveurs autorisés dans l'AD, étendues IPv4, état et taux
      d'occupation (les étendues > 90 % sont signalées).
    - DNS  : zones hébergées par les contrôleurs de domaine, type, mises à jour
      dynamiques, nombre d'enregistrements.
    Dates au format JJ-MM-AAAA.

.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-DHCP-DNSReport.ps1

.NOTES
    Prérequis : modules RSAT DhcpServer et/ou DnsServer + ActiveDirectory +
    ImportExcel (auto-installé). Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_DHCP_DNS_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
)
$ErrorActionPreference = 'Stop'
$Author = 'Valérian DUFOUR / Claude'
function Write-Step { param($m) Write-Host "[+] $m" -ForegroundColor Cyan }
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Step "Installation d'ImportExcel..."; try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop } catch { throw "Install-Module ImportExcel -Scope CurrentUser" }
}
Import-Module ImportExcel -ErrorAction Stop

# --- DHCP ---
$dhcpRows = New-Object System.Collections.Generic.List[object]
if (Get-Module -ListAvailable -Name DhcpServer) {
    Import-Module DhcpServer -ErrorAction SilentlyContinue
    Write-Step "Lecture des serveurs DHCP..."
    $servers = @()
    try { $servers = (Get-DhcpServerInDC).DnsName } catch { Write-Warning "Get-DhcpServerInDC : $($_.Exception.Message)" }
    foreach ($srv in $servers) {
        try {
            foreach ($sc in (Get-DhcpServerv4Scope -ComputerName $srv -ErrorAction Stop)) {
                $st = $null; try { $st = Get-DhcpServerv4ScopeStatistics -ComputerName $srv -ScopeId $sc.ScopeId -ErrorAction Stop } catch {}
                $pct = if ($st) { [math]::Round([double]$st.PercentageInUse,1) } else { $null }
                $dhcpRows.Add([PSCustomObject]@{
                    'Serveur'=$srv; 'Étendue'=$sc.ScopeId.IPAddressToString; 'Nom'=$sc.Name; 'État'="$($sc.State)"
                    'Occupation (%)'=$pct; 'Adresses utilisées'=(if($st){$st.AddressesInUse}else{$null}); 'Adresses libres'=(if($st){$st.AddressesFree}else{$null})
                    'Alerte'=(if($pct -ne $null -and $pct -ge 90){'Saturation (>90%)'}else{''})
                })
            }
        } catch { Write-Warning "DHCP $srv : $($_.Exception.Message)" }
    }
} else { Write-Warning "Module DhcpServer absent : partie DHCP ignorée." }

# --- DNS ---
$dnsRows = New-Object System.Collections.Generic.List[object]
if ((Get-Module -ListAvailable -Name DnsServer) -and (Get-Module -ListAvailable -Name ActiveDirectory)) {
    Import-Module DnsServer -ErrorAction SilentlyContinue; Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    Write-Step "Lecture des zones DNS..."
    $dcs = @(); try { $dcs = (Get-ADDomainController -Filter *).HostName } catch { $dcs = @($env:COMPUTERNAME) }
    $srv = $dcs | Select-Object -First 1
    try {
        foreach ($z in (Get-DnsServerZone -ComputerName $srv -ErrorAction Stop | Where-Object { -not $_.IsAutoCreated })) {
            $rc = $null; try { $rc = @(Get-DnsServerResourceRecord -ComputerName $srv -ZoneName $z.ZoneName -ErrorAction Stop).Count } catch {}
            $dnsRows.Add([PSCustomObject]@{
                'Serveur'=$srv; 'Zone'=$z.ZoneName; 'Type'="$($z.ZoneType)"; 'Intégrée AD'=(if($z.IsDsIntegrated){'Oui'}else{'Non'})
                'Mise à jour dynamique'="$($z.DynamicUpdate)"; 'Enregistrements'=$rc
            })
        }
    } catch { Write-Warning "DNS $srv : $($_.Exception.Message)" }
} else { Write-Warning "Module DnsServer/ActiveDirectory absent : partie DNS ignorée." }

$nbScopes=@($dhcpRows).Count; $nbSat=@($dhcpRows | Where-Object { $_.Alerte }).Count; $nbZones=@($dnsRows).Count
$nbSrvDhcp=@($dhcpRows.Serveur | Select-Object -Unique).Count
Write-Step ("DHCP: {0} étendue(s) ({1} saturée(s)) | DNS: {2} zone(s)." -f $nbScopes,$nbSat,$nbZones)

Write-Step "Génération Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$script:excel = $null
function Add-Detail { param($Data,$Sheet,$Table)
    if (@($Data).Count -eq 0) { $Data = ,([PSCustomObject]@{ Information='Aucun élément / module absent' }) }
    $script:excel = if ($null -eq $script:excel) { $Data | Export-Excel -Path $OutputPath -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
                    else { $Data | Export-Excel -ExcelPackage $script:excel -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
}
Add-Detail -Data $dhcpRows -Sheet 'DHCP étendues' -Table 'DHCP'
if (@($dhcpRows).Count -gt 0) {
    $ws=$script:excel.Workbook.Worksheets['DHCP étendues']; $hr=$ws.Dimension.Start.Row
    $rg = "A{0}:{1}{2}" -f ($hr+1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($ws.Dimension.End.Column), $ws.Dimension.End.Row
    # E = Occupation (%)
    Add-ConditionalFormatting -Worksheet $ws -Range $rg -RuleType Expression -ConditionValue ('=$E{0}>=90' -f ($hr+1)) -BackgroundColor ([System.Drawing.Color]::MistyRose) -Bold
}
Add-Detail -Data $dnsRows -Sheet 'DNS zones' -Table 'DNS'
$excel = $script:excel

$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'DHCP & DNS'
$wsK.Cells['B2'].Style.Font.Size = 16; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f (Get-Date -Format 'dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$kpis = @(
    @('Serveurs DHCP', $nbSrvDhcp, $null),
    @('Étendues DHCP', $nbScopes, $null),
    @('Étendues saturées (>90%)', $nbSat, 'red'),
    @('Zones DNS', $nbZones, $null)
)
$r = 7; $wsK.Cells["B$r"].Value='Indicateur'; $wsK.Cells["C$r"].Value='Valeur'
$wsK.Cells["B$r:C$r"].Style.Font.Bold=$true
$wsK.Cells["B$r:C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$r:C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$r:C$r"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) { $r++; $wsK.Cells["B$r"].Value=$k[0]; $wsK.Cells["C$r"].Value=$k[1]
    if ($k[2] -eq 'red' -and $k[1] -gt 0) { $wsK.Cells["C$r"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["C$r"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(255,199,206)); $wsK.Cells["C$r"].Style.Font.Bold=$true } }
$wsK.Column(2).Width=36; $wsK.Column(3).Width=16
Close-ExcelPackage $excel
Write-Host ("`n[OK] Rapport généré : {0}" -f $OutputPath) -ForegroundColor Green
