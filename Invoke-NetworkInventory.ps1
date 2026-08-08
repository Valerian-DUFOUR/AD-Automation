<#
.SYNOPSIS
    Inventaire reseau multi-sites + controle de conformite VLAN (generique).

.DESCRIPTION
    Balaye toutes les plages IP (sites x VLAN) du plan d'adressage configure,
    recupere l'adresse MAC de chaque equipement, resout le fabricant via la base
    OUI (IEEE), determine le TYPE d'equipement (PC, serveur, NAS, imprimante,
    borne WiFi Aruba/Ubiquiti, switch, badgeuse Kelio/Bodet, camera IP, telephone
    IP, automate, onduleur...), localise la machine (site + VLAN), puis verifie si
    elle se trouve dans le BON VLAN par rapport aux regles d'adressage definies.

    Trois moteurs de decouverte (combinables) :
      1) SNMP sur les passerelles (recommande) : lecture de la table ARP
         ipNetToMediaPhysAddress -> IP->MAC de TOUS les sous-reseaux routes, meme
         a distance. C'est LA methode fiable pour recuperer les MAC a travers les
         VLAN/sites (l'ARP local ne voit que le sous-reseau du scanner).
      2) Ping sweep + scan de ports TCP + banniere HTTP + reverse DNS (par hote).
      3) ARP local (Get-NetNeighbor / arp -a) pour le sous-reseau du scanner.

    Classification par MOTEUR DE SCORE : chaque signal (OUI, ports ouverts,
    banniere, hostname, SNMP) vote pour un type. Le VLAN d'appartenance n'est
    JAMAIS utilise pour decider du type (sinon on masquerait les non-conformites).

.PARAMETER Entreprise
    Nom de l'entreprise, affiche dans le titre des rapports.

.PARAMETER BaseOctet
    Premier octet du plan d'adressage (defaut 10 -> 10.<site>.<vlan>.<hote>).

.PARAMETER SitesFile
    Chemin d'un CSV (delimiteur ';') decrivant le plan de sites, colonnes :
    Octet;Societe;Ville;Id. S'il est fourni, il remplace la table interne.

.PARAMETER Sites
    Filtre : n'analyse que ces ID de site (2e octet). Ex : -Sites 1,20,30
    Par defaut : tous les sites connus.

.PARAMETER Vlans
    Filtre : n'analyse que ces VLAN (3e octet). Ex : -Vlans 1,2,17,50
    Par defaut : tous les VLAN definis dans le plan.

.PARAMETER UseSnmpGateways
    Active la moisson des tables ARP des passerelles via SNMP (fortement conseille
    pour recuperer les MAC a travers les sites/VLAN).

.PARAMETER Communaute
    Communaute SNMP en lecture (defaut : "public").

.PARAMETER UseSnmpHosts
    Interroge aussi chaque hote en SNMP (sysDescr/sysName). Desactive par defaut
    (SNMP souvent coupe sur les postes/terminaux du parc).

.PARAMETER OuiFile
    Chemin du fichier oui-db.csv (defaut : a cote du script). Sinon table interne.

.PARAMETER Diagnostic
    Analyse approfondie d'une seule IP (dump ports/banniere/SNMP + interpretation).

.EXAMPLE
    .\Invoke-NetworkInventory.ps1 -UseSnmpGateways
    .\Invoke-NetworkInventory.ps1 -Sites 20 -Vlans 1,2,17,50 -UseSnmpGateways
    .\Invoke-NetworkInventory.ps1 -Diagnostic 10.20.50.40

.NOTES
    Compatible PowerShell 5.1 et 7+. Aucun module obligatoire (Excel via ImportExcel
    si present). Lancer de preference depuis le controleur de domaine (ou un hote ayant une
    route vers tous les sites) pour maximiser la couverture reseau.
    Auteur : Valerian DUFOUR / Claude
#>

[CmdletBinding()]
param(
    [string] $Entreprise      = "Votre entreprise",   # nom affiche dans les rapports
    [int]    $BaseOctet       = 10,                    # 1er octet du plan (ex: 10 -> 10.site.vlan.hote)
    [string] $SitesFile,                              # CSV optionnel du plan de sites (Octet;Societe;Ville;Id)
    [int[]]  $Sites,
    [int[]]  $Vlans,
    [switch] $UseSnmpGateways,
    [string] $Communaute      = "public",
    [switch] $UseSnmpHosts,
    [switch] $ScanArpOnly,   # ne sonder que les IP presentes dans les tables ARP moissonnees (rapide)
    [string] $OuiFile,
    [int]    $HostStart       = 1,
    [int]    $HostEnd         = 254,
    [int]    $MaxThreads      = 128,
    [int]    $PingTimeoutMs   = 500,
    [int]    $TcpTimeoutMs    = 400,
    [int]    $HttpTimeoutSec  = 2,
    [int]    $SnmpTimeoutMs   = 800,
    [switch] $ResolveDns,
    [string] $Diagnostic
)

# =====================================================================
# ============================ CONFIGURATION ==========================
# =====================================================================

# ---- Sorties ----
$OutDir        = $PSScriptRoot
$FichierCSV    = Join-Path $OutDir "Inventaire_Reseau.csv"
$FichierJSON   = Join-Path $OutDir "Inventaire_Reseau.json"
$FichierExcel  = Join-Path $OutDir "Inventaire_Reseau.xlsx"
$FichierAnoCSV = Join-Path $OutDir "Inventaire_NonConformites.csv"

# ---- Ports TCP sondes (equilibre vitesse / signal) ----
$TcpPorts = @(22,23,80,135,139,443,445,515,554,631,2000,3389,5000,5001,5060,8080,8443,9100)

# ---- SITES : 2e octet -> etablissement (A PERSONNALISER ou via -SitesFile) ----
$SiteMap = @{
    1  = @{ Societe='Site principal (siege)'; Ville='Ville 1'; Id='SITE-01' }
    2  = @{ Societe='Site secondaire';        Ville='Ville 2'; Id='SITE-02' }
    3  = @{ Societe='Site distant';           Ville='Ville 3'; Id='SITE-03' }
}
# >>> PERSONNALISATION : editez la table ci-dessus (cle = 2e octet de l'IP),
#     OU fournissez -SitesFile <chemin.csv> (colonnes : Octet;Societe;Ville;Id).
if ($SitesFile -and (Test-Path $SitesFile)) {
    $SiteMap = @{}
    Import-Csv -Path $SitesFile -Delimiter ';' | ForEach-Object {
        if ($_.Octet) { $SiteMap[[int]$_.Octet] = @{ Societe=$_.Societe; Ville=$_.Ville; Id=$_.Id } }
    }
    Write-Host ("[SITES] Plan charge depuis {0} ({1} site(s))." -f $SitesFile, $SiteMap.Count) -ForegroundColor Green
}

# ---- VLAN : 3e octet -> usage + types d'equipements attendus ----
#  'Attendu' = liste des types normalement presents dans ce VLAN (pour la conformite).
$VlanMap = @{
    0   = @{ Nom='Serveurs / routeurs / firewalls';         Attendu=@('Serveur','Routeur','Firewall','Switch','NAS') }
    1   = @{ Nom='Postes DHCP dynamique';                   Attendu=@('PC') }
    2   = @{ Nom='Imprimantes (reservation DHCP)';          Attendu=@('Imprimante') }
    3   = @{ Nom='Libre';                                   Attendu=@() }
    4   = @{ Nom='DMZ Site';                                Attendu=@('Serveur','Firewall') }
    5   = @{ Nom='VM Archives Site';                        Attendu=@('Serveur') }
    15  = @{ Nom='Admin materiels (ESX,iDRAC,onduleurs,baies,APC)'; Attendu=@('Onduleur','Serveur','NAS','Switch','AdminMat') }
    16  = @{ Nom='Indus : infra switchs/routeurs/fw';       Attendu=@('Switch','Routeur','Firewall') }
    17  = @{ Nom='Indus : postes PC';                       Attendu=@('PC') }
    18  = @{ Nom='Indus : automates/capteurs/MES';          Attendu=@('Automate','Capteur','PC') }
    19  = @{ Nom='Indus : libre';                           Attendu=@() }
    20  = @{ Nom='Serveurs';                                Attendu=@('Serveur','NAS') }
    30  = @{ Nom='VLAN de test';                      Attendu=@() }
    32  = @{ Nom='WiFi : infrastructure (bornes,switchs)';  Attendu=@('BorneWiFi','Switch','Routeur') }
    33  = @{ Nom='WiFi administratif';                      Attendu=@('PC','Smartphone') }
    34  = @{ Nom='WiFi industriel';                         Attendu=@('PC','Automate') }
    35  = @{ Nom='WiFi invites';                            Attendu=@('PC','Smartphone') }
    36  = @{ Nom='WiFi mobiles internes (smartphones/tablettes)'; Attendu=@('Smartphone') }
    37  = @{ Nom='WiFi reserve';                            Attendu=@() }
    38  = @{ Nom='WiFi reserve';                            Attendu=@() }
    39  = @{ Nom='WiFi reserve';                            Attendu=@() }
    40  = @{ Nom='Quarantaine';                             Attendu=@() }
    48  = @{ Nom='Videoprotection (cameras IP, enregistreurs)'; Attendu=@('CameraIP','NVR') }
    49  = @{ Nom='Alarme';                                  Attendu=@('Alarme') }
    50  = @{ Nom='GTC/GTB (Badgeuse Kelio, clim, chauffage)'; Attendu=@('Badgeuse','GTC') }
    51  = @{ Nom='Bornes recharge vehicules electriques';   Attendu=@('BorneVE') }
    52  = @{ Nom='VPN Indus (boitiers prestataires)';       Attendu=@('VPN') }
    53  = @{ Nom='Videoprotection production';         Attendu=@('CameraIP','NVR') }
    64  = @{ Nom='Telephonie IP : infra (IPPBX, passerelles)'; Attendu=@('TelephoneIP','Routeur') }
    65  = @{ Nom='Telephones IP (DHCP dynamique)';          Attendu=@('TelephoneIP') }
    66  = @{ Nom='Telephonie : libre';                      Attendu=@() }
    67  = @{ Nom='Telephonie : libre';                      Attendu=@() }
    100 = @{ Nom='Reseau automates/PLC (liaison directe firewall)'; Attendu=@('Automate','PC') }
    101 = @{ Nom='Reseau automates/PLC';                    Attendu=@('Automate') }
    102 = @{ Nom='Reseau automates/PLC';                    Attendu=@('Automate') }
    103 = @{ Nom='Reseau automates/PLC';                    Attendu=@('Automate') }
}

# ---- VLAN cibles par TYPE d'equipement (ou il DEVRAIT etre) ----
#  Premier = VLAN nominal. Les suivants = toleres. Sert au calcul de conformite.
$TypeToVlan = @{
    'PC'          = @(1,17,33,34,35)     # bureautique .1, indus .17, WiFi .33/34/35
    'Serveur'     = @(0,20,4,5)
    'NAS'         = @(20,0,15)
    'Imprimante'  = @(2)
    'BorneWiFi'   = @(32)
    'Switch'      = @(0,16,32,15)
    'Routeur'     = @(0,16,64,32)
    'Firewall'    = @(0,16,4)
    'Onduleur'    = @(15)
    'AdminMat'    = @(15)
    'CameraIP'    = @(48,53)
    'NVR'         = @(48,53)
    'Badgeuse'    = @(50)
    'GTC'         = @(50)
    'TelephoneIP' = @(64,65)
    'Automate'    = @(18,100,101,102,103,16)
    'Capteur'     = @(18,100)
    'BorneVE'     = @(51)
    'Alarme'      = @(49)
    'Smartphone'  = @(36,33,35)
    'VPN'         = @(52)
    'Terminal'    = @(1,17)               # clients legers Axel
}

# ---- Passerelles a interroger en SNMP (par site) pour moissonner les MAC ----
#  Chaque passerelle detient la table ARP des sous-reseaux qu'elle route.
$GatewayHostOctets = @(0,4,5,15,16,20,48,49,50,51,64)   # .254 sur chacun

# ---- Table OUI interne de repli (si oui-db.csv absent) ----
#  Cle = 3 premiers octets MAC ; valeur = fabricant. La base CSV externe (38k+
#  entrees) prime si presente. Ici : fabricants courants (table de repli generique).
$OuiFallback = @{
    '00:A0:34'='Axel';                 'EC:9B:8B'='Hewlett Packard Enterprise';
    '00:0B:84'='Bodet';                '14:18:77'='Dell';
    'F4:8E:38'='Dell';                 '94:2A:6F'='Ubiquiti';
    'F4:92:BF'='Ubiquiti';             'B4:FB:E4'='Ubiquiti';
    '68:D7:9A'='Ubiquiti';             '74:83:C2'='Ubiquiti';
    'AC:8B:A9'='Ubiquiti';             '78:45:58'='Ubiquiti';
    '44:91:60'='Murata Manufacturing'; '00:E0:4D'='Internet Initiative Japan';
    '08:3A:2F'='Guangzhou Juan Intelligent';
}

# ---- Regles de classification par fabricant (mot-cle dans le nom OUI -> type) ----
$VendorTypeRules = @(
    @{ Match='axel';                                    Type='Terminal';    Poids=6 }
    @{ Match='zebra|sato|toshiba tec|tsc|godex|citizen';Type='Imprimante';  Poids=5 }
    @{ Match='brother|kyocera|ricoh|konica|lexmark|canon|epson|xerox|sharp|oki|develop';Type='Imprimante';Poids=4 }
    @{ Match='aruba|hewlett packard enterprise.*aruba|ruckus|extreme network|mist';Type='BorneWiFi';Poids=4 }
    @{ Match='ubiquiti';                                Type='BorneWiFi';   Poids=3 }
    @{ Match='cisco meraki|meraki';                     Type='BorneWiFi';   Poids=4 }
    @{ Match='sonicwall|fortinet|palo alto|watchguard|stormshield|checkpoint';Type='Firewall';Poids=6 }
    @{ Match='cisco|juniper|arista|netgear|d-link|tp-link|mikrotik|hp.*procurve|allied telesis|h3c|huawei';Type='Switch';Poids=3 }
    @{ Match='synology|qnap|western digital|buffalo|netapp|drobo|terra-master|asustor';Type='NAS';Poids=5 }
    @{ Match='apc|american power|schneider.*apc|eaton|riello|legrand|vertiv|liebert|tripp|socomec|cyberpower';Type='Onduleur';Poids=5 }
    @{ Match='bodet';                                   Type='Badgeuse';    Poids=5 }
    @{ Match='axis|hikvision|hangzhou hikvision|dahua|hanwha|mobotix|bosch.*security|avigilon|vivotek|uniview|milestone';Type='CameraIP';Poids=5 }
    @{ Match='yealink|polycom|snom|grandstream|mitel|avaya|gigaset|alcatel-lucent enterprise|fanvil|htek|sangoma';Type='TelephoneIP';Poids=5 }
    @{ Match='siemens|schneider electric|rockwell|allen-bradley|beckhoff|wago|phoenix contact|omron|mitsubishi electric|b&r|hilscher|pilz|turck|ifm|balluff';Type='Automate';Poids=5 }
    @{ Match='vmware|nutanix';                          Type='Serveur';     Poids=4 }
    @{ Match='super micro|supermicro|quanta|inspur|gigabyte|tyan|asrock rack';Type='Serveur';Poids=3 }
    @{ Match='apple';                                   Type='Smartphone';  Poids=2 }
    @{ Match='samsung|xiaomi|oneplus|google.*pixel|huawei device|honor';Type='Smartphone';Poids=2 }
    @{ Match='dell|lenovo|hp inc|hewlett.packard(?!.*enterprise)|fujitsu|acer|asustek|intel corporate|micro-star|msi|toshiba client';Type='PC';Poids=2 }
    @{ Match='raspberry|espressif|texas instruments|advantech|moxa';Type='Automate';Poids=2 }
)

# =====================================================================
# ======================= FONCTIONS UTILITAIRES =======================
# =====================================================================

# --- Bloc de helpers partages (injecte dans les runspaces ET dot-source ici) ---
$SharedHelpers = @'
function Format-Mac([string]$raw) {
    if (-not $raw) { return "" }
    $hex = ($raw -replace '[^0-9A-Fa-f]','').ToUpper()
    if ($hex.Length -lt 12) { return "" }
    $hex = $hex.Substring(0,12)
    $p = for ($i=0;$i -lt 12;$i+=2){ $hex.Substring($i,2) }
    return ($p -join ':')
}
function Get-Oui([string]$mac) {
    if (-not $mac) { return "" }
    $h = ($mac -replace '[^0-9A-Fa-f]','').ToUpper()
    if ($h.Length -lt 6) { return "" }
    return "{0}:{1}:{2}" -f $h.Substring(0,2),$h.Substring(2,2),$h.Substring(4,2)
}
function Test-TcpPort([string]$Ip,[int]$Port,[int]$TimeoutMs) {
    $c = New-Object System.Net.Sockets.TcpClient
    try {
        $iar = $c.BeginConnect($Ip,$Port,$null,$null)
        if ($iar.AsyncWaitHandle.WaitOne($TimeoutMs,$false) -and $c.Connected) {
            $c.EndConnect($iar); return $true
        }
        return $false
    } catch { return $false } finally { $c.Close() }
}
function Get-HttpBanner([string]$Scheme,[string]$Ip,[int]$Port,[int]$TimeoutSec,[int]$PSMajor) {
    $url = "{0}://{1}:{2}/" -f $Scheme,$Ip,$Port
    try {
        try { [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12 -bor [Net.SecurityProtocolType]::Tls11 -bor [Net.SecurityProtocolType]::Tls } catch {}
        if ($PSMajor -ge 6) {
            $r = Invoke-WebRequest -Uri $url -TimeoutSec $TimeoutSec -UseBasicParsing -SkipCertificateCheck -MaximumRedirection 1 -ErrorAction Stop
        } else {
            [Net.ServicePointManager]::ServerCertificateValidationCallback = { $true }
            $r = Invoke-WebRequest -Uri $url -TimeoutSec $TimeoutSec -UseBasicParsing -MaximumRedirection 1 -ErrorAction Stop
        }
        $srv = ""; try { $srv = [string]$r.Headers['Server'] } catch {}
        $auth= ""; try { $auth= [string]$r.Headers['WWW-Authenticate'] } catch {}
        $body= [string]$r.Content; if ($body.Length -gt 4000) { $body = $body.Substring(0,4000) }
        $title = ""
        $m = [regex]::Match($body,'<title[^>]*>(.*?)</title>','IgnoreCase, Singleline')
        if ($m.Success) { $title = ($m.Groups[1].Value -replace '\s+',' ').Trim() }
        return [PSCustomObject]@{ Server=$srv; Title=$title; Auth=$auth }
    } catch { return $null }
}
'@
Invoke-Expression $SharedHelpers

# ---------------- Mini-pile SNMP (BER) pour GET + WALK ARP -------------
$SnmpHelpers = @'
function Get-BerLen([int]$n){ if($n -lt 0x80){return ,([byte]$n)}; $b=New-Object System.Collections.Generic.List[byte]; $v=$n; while($v -gt 0){$b.Insert(0,[byte]($v -band 0xFF)); $v=$v -shr 8}; $o=New-Object System.Collections.Generic.List[byte]; $o.Add([byte](0x80 -bor $b.Count)); $o.AddRange($b); return $o.ToArray() }
function New-Tlv([byte]$t,[byte[]]$c){ if($null -eq $c){$c=[byte[]]@()}; $r=New-Object System.Collections.Generic.List[byte]; $r.Add($t); $r.AddRange((Get-BerLen $c.Length)); if($c.Length){$r.AddRange($c)}; return $r.ToArray() }
function Get-OidB([string]$oid){ $p=$oid.Split('.')|ForEach-Object{[int]$_}; $b=New-Object System.Collections.Generic.List[byte]; $b.Add([byte](40*$p[0]+$p[1])); for($i=2;$i -lt $p.Count;$i++){ $v=$p[$i]; if($v -lt 0x80){$b.Add([byte]$v)} else { $s=New-Object System.Collections.Generic.List[byte]; $s.Add([byte]($v -band 0x7F)); $v=$v -shr 7; while($v -gt 0){$s.Add([byte](($v -band 0x7F)-bor 0x80)); $v=$v -shr 7}; $s.Reverse(); $b.AddRange($s) } }; return $b.ToArray() }
function Get-OidS([byte[]]$b){ if(-not $b -or $b.Length -eq 0){return ""}; $f=[int][math]::Floor($b[0]/40); $sec=$b[0]%40; if($f -gt 2){$f=2;$sec=$b[0]-80}; $a=New-Object System.Collections.Generic.List[string]; $a.Add("$f"); $a.Add("$sec"); $val=0; for($i=1;$i -lt $b.Length;$i++){ $val=($val -shl 7)-bor($b[$i]-band 0x7F); if(-not($b[$i]-band 0x80)){$a.Add("$val");$val=0} }; return ($a -join '.') }
function Read-Tlv([byte[]]$d,[int]$o){ $t=$d[$o];$o++; $l=$d[$o];$o++; if($l -band 0x80){ $nb=$l -band 0x7F;$l=0; for($k=0;$k -lt $nb;$k++){$l=($l -shl 8)-bor $d[$o];$o++} }; return [PSCustomObject]@{Tag=$t;Start=$o;Len=$l;Next=($o+$l)} }
function Invoke-SnmpRaw([string]$Ip,[string]$Comm,[string]$Oid,[int]$TimeoutMs,[byte]$Pdu){
    try {
        $ver=New-Tlv 0x02 ([byte[]]@(1)); $cm=New-Tlv 0x04 ([Text.Encoding]::ASCII.GetBytes($Comm))
        $rid=[byte[]]::new(4); (New-Object Random).NextBytes($rid); $rid[0]=$rid[0]-band 0x7F
        $ri=New-Tlv 0x02 $rid; $e1=New-Tlv 0x02 ([byte[]]@(0)); $e2=New-Tlv 0x02 ([byte[]]@(0))
        $ot=New-Tlv 0x06 (Get-OidB $Oid); $nl=New-Tlv 0x05 ([byte[]]@())
        $vb=New-Tlv 0x30 ([byte[]]($ot+$nl)); $vl=New-Tlv 0x30 $vb
        $pd=New-Tlv $Pdu ([byte[]]($ri+$e1+$e2+$vl)); $msg=New-Tlv 0x30 ([byte[]]($ver+$cm+$pd))
        $u=New-Object Net.Sockets.UdpClient; $u.Client.ReceiveTimeout=$TimeoutMs; $u.Connect($Ip,161)
        [void]$u.Send($msg,$msg.Length); $rep=New-Object Net.IPEndPoint([Net.IPAddress]::Any,0)
        $resp=$u.Receive([ref]$rep); $u.Close()
        $t=Read-Tlv $resp 0;$o=$t.Start; $t=Read-Tlv $resp $o;$o=$t.Next; $t=Read-Tlv $resp $o;$o=$t.Next
        $t=Read-Tlv $resp $o;$o=$t.Start; $t=Read-Tlv $resp $o;$o=$t.Next; $t=Read-Tlv $resp $o;$o=$t.Next
        $t=Read-Tlv $resp $o;$o=$t.Next; $t=Read-Tlv $resp $o;$o=$t.Start; $t=Read-Tlv $resp $o;$o=$t.Start
        $ot2=Read-Tlv $resp $o
        $rOid=Get-OidS ([byte[]]($resp[$ot2.Start..($ot2.Start+$ot2.Len-1)]))
        $vt=Read-Tlv $resp $ot2.Next; $tag=$vt.Tag
        if($vt.Len -gt 0){ $raw=[byte[]]($resp[$vt.Start..($vt.Start+$vt.Len-1)]) } else { $raw=[byte[]]@() }
        $hex=""; if($raw.Length){ $hex=($raw|ForEach-Object{$_.ToString('X2')}) -join ':' }
        $str=""; try { $str=[Text.Encoding]::ASCII.GetString($raw) } catch {}
        return [PSCustomObject]@{ OID=$rOid; Tag=$tag; Hex=$hex; Str=$str }
    } catch { return $null }
}
function Invoke-SnmpWalkArp([string]$Ip,[string]$Comm,[int]$TimeoutMs,[int]$Max){
    # Walk ipNetToMediaPhysAddress = 1.3.6.1.2.1.4.22.1.2 -> IP->MAC
    $base='1.3.6.1.2.1.4.22.1.2'; $cur=$base
    $res=New-Object System.Collections.Generic.List[object]
    for($n=0;$n -lt $Max;$n++){
        $r=Invoke-SnmpRaw $Ip $Comm $cur $TimeoutMs 0xA1
        if(-not $r){break}
        if($r.Tag -eq 0x80 -or $r.Tag -eq 0x81 -or $r.Tag -eq 0x82){break}
        if(-not ($r.OID -eq $base -or $r.OID.StartsWith($base+'.'))){break}
        if($r.OID -eq $cur){break}
        # OID suffixe : .<ifIndex>.<a>.<b>.<c>.<d>  -> les 4 derniers arcs = IP
        $arcs=$r.OID.Split('.'); if($arcs.Count -ge 4){ $ip4=($arcs[-4..-1]) -join '.' } else { $ip4="" }
        if($ip4 -and $r.Hex){ $res.Add([PSCustomObject]@{ IP=$ip4; MacHex=$r.Hex }) }
        $cur=$r.OID
    }
    return $res
}
'@
Invoke-Expression $SnmpHelpers

# ---------------- Chargement base OUI ----------------
function Import-OuiDatabase {
    param([string]$Path)
    $table = @{}
    if (-not $Path) { $Path = Join-Path $PSScriptRoot 'oui-db.csv' }
    if (Test-Path $Path) {
        try {
            $rows = Import-Csv -Path $Path -Delimiter ';'
            foreach ($r in $rows) {
                $k = ([string]$r.OUI).ToUpper()
                if ($k) { $table[$k] = [string]$r.Vendor }
            }
            Write-Host ("[OUI] Base externe chargee : {0} entrees ({1})" -f $table.Count, (Split-Path $Path -Leaf)) -ForegroundColor Green
        } catch { Write-Host "[OUI] Echec lecture $Path : $_" -ForegroundColor Yellow }
    }
    if ($table.Count -eq 0) {
        foreach ($k in $OuiFallback.Keys) { $table[$k] = $OuiFallback[$k] }
        Write-Host ("[OUI] Table interne de repli : {0} entrees (placez oui-db.csv a cote du script pour +38000)" -f $table.Count) -ForegroundColor Yellow
    }
    return $table
}

# ---------------- Resolution site / vlan ----------------
function Resolve-Localisation {
    param([string]$Ip)
    $o = $Ip.Split('.')
    $siteOct = [int]$o[1]; $vlanOct = [int]$o[2]
    $site = $SiteMap[$siteOct]
    $vlan = $VlanMap[$vlanOct]
    return [PSCustomObject]@{
        SiteOctet = $siteOct
        VlanOctet = $vlanOct
        Societe   = $(if ($site){$site.Societe}else{"Site inconnu ($siteOct)"})
        Ville     = $(if ($site){$site.Ville}else{"?"})
        SiteId    = $(if ($site){$site.Id}else{"?"})
        VlanNom   = $(if ($vlan){$vlan.Nom}else{"VLAN non defini ($vlanOct)"})
        VlanAttendu = $(if ($vlan){$vlan.Attendu}else{@()})
    }
}

# ---------------- Moteur de classification ----------------
function Get-DeviceType {
    param(
        [string]$Vendor,[int[]]$Ports,[string]$HttpServer,[string]$HttpTitle,
        [string]$HttpAuth,[string]$Hostname,[string]$SnmpDescr
    )
    $scores = @{}
    $why = New-Object System.Collections.Generic.List[string]
    function Add-Score($t,$w,$r){ if(-not $scores.ContainsKey($t)){$scores[$t]=0}; $scores[$t]+=$w; $why.Add("$r(+$w->$t)") }

    # 1) Fabricant (OUI)
    if ($Vendor) {
        $vl = $Vendor.ToLower()
        foreach ($rule in $VendorTypeRules) {
            if ($vl -match $rule.Match) { Add-Score $rule.Type $rule.Poids "OUI:$Vendor" ; break }
        }
    }
    # 2) Ports ouverts
    $p = @{}; foreach($x in $Ports){$p[$x]=$true}
    if ($p[9100] -or $p[515] -or $p[631]) { Add-Score 'Imprimante' 4 'port9100/515/631' }
    if ($p[554])                          { Add-Score 'CameraIP'  3 'port554-RTSP' }
    if ($p[5060])                         { Add-Score 'TelephoneIP' 3 'port5060-SIP' }
    if ($p[5000] -or $p[5001])            { Add-Score 'NAS'       2 'port5000/5001' }
    if ($p[445] -and $p[3389])            { Add-Score 'Serveur'   1 'SMB+RDP' }
    if ($p[3389] -and -not $p[445])       { Add-Score 'PC'        1 'RDP' }
    if ($p[445] -and -not $p[3389] -and -not $p[80] -and -not $p[443]) { Add-Score 'PC' 1 'SMB' }
    if ($p[22] -and -not $p[445])         { Add-Score 'Serveur'   1 'SSH' }
    # 3) Banniere HTTP / titre / auth
    $ban = ("{0} {1} {2} {3}" -f $HttpServer,$HttpTitle,$HttpAuth,$SnmpDescr).ToLower()
    if ($ban) {
        if ($ban -match 'aruba|instant|arubaos')          { Add-Score 'BorneWiFi' 4 'ban-aruba' }
        if ($ban -match 'unifi|ubnt|edgeos|edgeswitch')    { Add-Score 'BorneWiFi' 3 'ban-ubnt' }
        if ($ban -match 'sonicwall|fortigate|pfsense|stormshield|watchguard') { Add-Score 'Firewall' 5 'ban-fw' }
        if ($ban -match 'cisco|procurve|powerconnect|dell networking|switch') { Add-Score 'Switch' 3 'ban-switch' }
        if ($ban -match 'synology|diskstation|qnap|truenas|freenas') { Add-Score 'NAS' 5 'ban-nas' }
        if ($ban -match 'idrac|ilo|integrated lights|cimc|redfish|xclarity|imm') { Add-Score 'AdminMat' 5 'ban-bmc' }
        if ($ban -match 'apc|smart-ups|network management card|eaton|riello|ups') { Add-Score 'Onduleur' 4 'ban-ups' }
        if ($ban -match 'axis|hikvision|dahua|web service|network camera|rtsp|milesight|reolink') { Add-Score 'CameraIP' 4 'ban-cam' }
        if ($ban -match 'bodet|kelio|sigma')               { Add-Score 'Badgeuse' 5 'ban-bodet' }
        if ($ban -match 'yealink|polycom|grandstream|snom|sip|voip|phone') { Add-Score 'TelephoneIP' 3 'ban-voip' }
        if ($ban -match 'jetdirect|hp laserjet|kyocera|ricoh|lexmark|xerox|printer|imprimante|brother|canon ij|epson') { Add-Score 'Imprimante' 4 'ban-print' }
        if ($ban -match 'vmware|esxi|proxmox|hyper-v|xenserver|idrac|poweredge|proliant') { Add-Score 'Serveur' 3 'ban-server' }
        if ($ban -match 'simatic|siemens|codesys|modbus|profinet|automation|plc|wago|beckhoff|schneider') { Add-Score 'Automate' 4 'ban-plc' }
        if ($ban -match 'axel|ax3000')                     { Add-Score 'Terminal' 5 'ban-axel' }
    }
    # 4) Hostname
    if ($Hostname) {
        $hn = $Hostname.ToLower()
        if ($hn -match 'print|impr|hpljet|kyocera|mfp')    { Add-Score 'Imprimante' 3 'hn-print' }
        if ($hn -match 'srv|server|dc0|dc1|vcenter|esx|host') { Add-Score 'Serveur' 2 'hn-srv' }
        if ($hn -match 'nas|synology|qnap|diskstation')    { Add-Score 'NAS' 3 'hn-nas' }
        if ($hn -match 'cam|camera|nvr|dvr')               { Add-Score 'CameraIP' 3 'hn-cam' }
        if ($hn -match 'ap-|wifi|aruba|unifi|borne')       { Add-Score 'BorneWiFi' 3 'hn-ap' }
        if ($hn -match 'sw-|switch|core|stack')            { Add-Score 'Switch' 2 'hn-sw' }
        if ($hn -match 'kelio|badge|pointeuse|gtc|gtb')    { Add-Score 'Badgeuse' 3 'hn-badge' }
        if ($hn -match 'phone|tel|voip|sip')               { Add-Score 'TelephoneIP' 2 'hn-tel' }
        if ($hn -match 'pc-|desktop|laptop|poste|wks|ws-') { Add-Score 'PC' 2 'hn-pc' }
        if ($hn -match 'plc|automate|api-|scada')          { Add-Score 'Automate' 3 'hn-plc' }
        if ($hn -match 'ups|onduleur')                     { Add-Score 'Onduleur' 3 'hn-ups' }
    }

    if ($scores.Count -eq 0) {
        return [PSCustomObject]@{ Type='Indetermine'; Confiance=0; Details='aucun signal' }
    }
    $best = $scores.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1
    return [PSCustomObject]@{
        Type      = $best.Key
        Confiance = $best.Value
        Details   = ($why -join '; ')
    }
}

# ---------------- Moteur de conformite VLAN ----------------
function Test-Conformite {
    param([string]$Type,[int]$VlanOctet,[object]$Loc)
    if ($Type -eq 'Indetermine') {
        return [PSCustomObject]@{ Statut='A VERIFIER'; VlanCible='?'; Message='Type non identifie' }
    }
    $cibles = $TypeToVlan[$Type]
    if (-not $cibles) {
        return [PSCustomObject]@{ Statut='A VERIFIER'; VlanCible='?'; Message="Pas de regle VLAN pour type '$Type'" }
    }
    # Correspondance directe VLAN cible
    if ($cibles -contains $VlanOctet) {
        $rang = [array]::IndexOf($cibles,$VlanOctet)
        if ($rang -eq 0) {
            return [PSCustomObject]@{ Statut='CONFORME'; VlanCible=($cibles -join '/'); Message="VLAN nominal" }
        } else {
            return [PSCustomObject]@{ Statut='TOLERE'; VlanCible=($cibles -join '/'); Message="VLAN tolere (nominal=.$($cibles[0]))" }
        }
    }
    # Le VLAN actuel accepte-t-il ce type (vue cote plan d'adressage) ?
    if ($Loc.VlanAttendu -and ($Loc.VlanAttendu -contains $Type)) {
        return [PSCustomObject]@{ Statut='TOLERE'; VlanCible=($cibles -join '/'); Message="Accepte par l'usage du VLAN" }
    }
    return [PSCustomObject]@{
        Statut    = 'NON CONFORME'
        VlanCible = ($cibles -join '/')
        Message   = "$Type detecte en .$VlanOctet -> attendu VLAN .$($cibles[0])"
    }
}

# =====================================================================
# ==================== MOISSON MAC VIA PASSERELLES ====================
# =====================================================================
function Invoke-GatewayMacHarvest {
    param([int[]]$SiteList,[string]$Comm,[int]$TimeoutMs)
    $macTable = @{}   # IP -> MAC formatee
    $gws = New-Object System.Collections.Generic.List[string]
    foreach ($s in $SiteList) {
        foreach ($ho in $GatewayHostOctets) { $gws.Add("$BaseOctet.$s.$ho.254") }
    }
    Write-Host ("[SNMP] Moisson ARP sur {0} passerelles potentielles..." -f $gws.Count) -ForegroundColor Cyan
    $ok = 0
    foreach ($gw in $gws) {
        # ping court avant SNMP pour ne pas perdre de temps
        try { $pr = ([Net.NetworkInformation.Ping]::new()).Send($gw,300); if ($pr.Status -ne 'Success') { continue } } catch { continue }
        $arp = Invoke-SnmpWalkArp $gw $Comm $TimeoutMs 2000
        if ($arp -and $arp.Count -gt 0) {
            $ok++
            foreach ($e in $arp) {
                $m = Format-Mac $e.MacHex
                if ($m -and -not $macTable.ContainsKey($e.IP)) { $macTable[$e.IP] = $m }
            }
            Write-Host ("   [+] {0,-16} : {1} entrees ARP" -f $gw,$arp.Count) -ForegroundColor DarkGreen
        }
    }
    Write-Host ("[SNMP] {0} passerelle(s) repondante(s), {1} couples IP/MAC collectes" -f $ok,$macTable.Count) -ForegroundColor Green
    return $macTable
}

# ---------------- ARP local (sous-reseau du scanner) ----------------
function Get-LocalArpTable {
    $t = @{}
    try {
        if (Get-Command Get-NetNeighbor -ErrorAction SilentlyContinue) {
            Get-NetNeighbor -AddressFamily IPv4 -ErrorAction SilentlyContinue |
                Where-Object { $_.State -in @('Reachable','Stale','Permanent') -and $_.LinkLayerAddress -match '..-..-..-..-..-..' } |
                ForEach-Object { $t[$_.IPAddress] = (Format-Mac $_.LinkLayerAddress) }
        } else {
            foreach ($line in (arp -a)) {
                $m = [regex]::Match($line,'(\d{1,3}(?:\.\d{1,3}){3})\s+((?:[0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2})')
                if ($m.Success) { $t[$m.Groups[1].Value] = (Format-Mac $m.Groups[2].Value) }
            }
        }
    } catch {}
    return $t
}

# =====================================================================
# ========================= MODE DIAGNOSTIC ===========================
# =====================================================================
if ($Diagnostic) {
    $ip = $Diagnostic
    $PSMajor = $PSVersionTable.PSVersion.Major
    $OuiTable = Import-OuiDatabase -Path $OuiFile
    Write-Host "=========== DIAGNOSTIC $ip ===========" -ForegroundColor Cyan
    $loc = Resolve-Localisation $ip
    Write-Host ("Site   : {0} - {1} ({2})  [octet {3}]" -f $loc.Societe,$loc.Ville,$loc.SiteId,$loc.SiteOctet)
    Write-Host ("VLAN   : .{0} - {1}" -f $loc.VlanOctet,$loc.VlanNom)
    try { $pr = ([Net.NetworkInformation.Ping]::new()).Send($ip,1500); Write-Host ("Ping   : {0}" -f $pr.Status) -ForegroundColor Yellow } catch { Write-Host "Ping   : echec" -ForegroundColor Red }

    Write-Host "`n--- Ports TCP ---" -ForegroundColor Cyan
    $open = @()
    foreach ($pt in $TcpPorts) { if (Test-TcpPort $ip $pt $TcpTimeoutMs) { $open += $pt; Write-Host ("   {0} ouvert" -f $pt) -ForegroundColor Green } }
    if (-not $open) { Write-Host "   (aucun port ouvert dans la liste)" -ForegroundColor DarkGray }

    Write-Host "`n--- Banniere HTTP/HTTPS ---" -ForegroundColor Cyan
    $b = $null
    foreach ($sc in @(@{S='http';P=80},@{S='https';P=443},@{S='http';P=8080},@{S='https';P=8443})) {
        if ($open -contains $sc.P) { $b = Get-HttpBanner $sc.S $ip $sc.P $HttpTimeoutSec $PSMajor; if ($b) { break } }
    }
    if ($b) { Write-Host ("   Server: {0}`n   Title : {1}`n   Auth  : {2}" -f $b.Server,$b.Title,$b.Auth) } else { Write-Host "   (pas de reponse web)" -ForegroundColor DarkGray }

    $mac = ""
    if ($UseSnmpGateways) {
        Write-Host "`n--- MAC via passerelle SNMP ---" -ForegroundColor Cyan
        $gw = "$BaseOctet.$($loc.SiteOctet).0.254"
        $arp = Invoke-SnmpWalkArp $gw $Communaute $SnmpTimeoutMs 2000
        $hit = $arp | Where-Object { $_.IP -eq $ip } | Select-Object -First 1
        if ($hit) { $mac = Format-Mac $hit.MacHex; Write-Host ("   $mac (via $gw)") -ForegroundColor Green } else { Write-Host "   (non trouvee sur $gw)" -ForegroundColor DarkGray }
    }
    if (-not $mac) { $la = Get-LocalArpTable; if ($la.ContainsKey($ip)) { $mac = $la[$ip]; Write-Host ("   MAC (ARP local) : $mac") -ForegroundColor Green } }

    $vendor = ""; if ($mac) { $oui = Get-Oui $mac; if ($OuiTable.ContainsKey($oui)) { $vendor = $OuiTable[$oui] } }
    $sd = ""; $sn = ""
    if ($UseSnmpHosts) { $r = Invoke-SnmpRaw $ip $Communaute '1.3.6.1.2.1.1.1.0' $SnmpTimeoutMs 0xA0; if ($r) { $sd = $r.Str } }
    $cls = Get-DeviceType -Vendor $vendor -Ports $open -HttpServer $(if($b){$b.Server}) -HttpTitle $(if($b){$b.Title}) -HttpAuth $(if($b){$b.Auth}) -Hostname "" -SnmpDescr $sd
    $conf = Test-Conformite -Type $cls.Type -VlanOctet $loc.VlanOctet -Loc $loc

    Write-Host "`n--- SYNTHESE ---" -ForegroundColor Cyan
    Write-Host ("   MAC        : {0}" -f $(if($mac){$mac}else{'(non recuperee)'}))
    Write-Host ("   Fabricant  : {0}" -f $(if($vendor){$vendor}else{'?'}))
    Write-Host ("   Type       : {0} (confiance {1})" -f $cls.Type,$cls.Confiance) -ForegroundColor Yellow
    Write-Host ("   Signaux    : {0}" -f $cls.Details) -ForegroundColor DarkGray
    $col = switch ($conf.Statut) { 'CONFORME'{'Green'} 'TOLERE'{'Cyan'} 'NON CONFORME'{'Red'} default{'Yellow'} }
    Write-Host ("   Conformite : {0} - {1}" -f $conf.Statut,$conf.Message) -ForegroundColor $col
    return
}

# =====================================================================
# ============================ SCAN COMPLET ===========================
# =====================================================================
$PSMajor  = $PSVersionTable.PSVersion.Major
$OuiTable = Import-OuiDatabase -Path $OuiFile

# --- Perimetre ---
$SiteList = if ($Sites) { $Sites } else { $SiteMap.Keys | Sort-Object }
$VlanList = if ($Vlans) { $Vlans } else { $VlanMap.Keys  | Sort-Object }

Write-Host ("`nPerimetre : {0} site(s) x {1} VLAN x {2} hotes" -f @($SiteList).Count,@($VlanList).Count,($HostEnd-$HostStart+1)) -ForegroundColor Cyan

# --- Table MAC (passerelles SNMP + ARP local) ---
$MacTable = @{}
if ($UseSnmpGateways) {
    $harvest = Invoke-GatewayMacHarvest -SiteList $SiteList -Comm $Communaute -TimeoutMs $SnmpTimeoutMs
    foreach ($k in $harvest.Keys) { $MacTable[$k] = $harvest[$k] }
} else {
    Write-Host "[INFO] -UseSnmpGateways non active : les MAC distantes ne seront pas recuperees (ARP local uniquement)." -ForegroundColor Yellow
}
$localArp = Get-LocalArpTable
foreach ($k in $localArp.Keys) { if (-not $MacTable.ContainsKey($k)) { $MacTable[$k] = $localArp[$k] } }

# --- Liste des IP a sonder ---
Write-Host "Generation de la liste d'IP..." -ForegroundColor Cyan
$AllIPs = [System.Collections.Generic.List[string]]::new()
if ($ScanArpOnly -and $MacTable.Count -gt 0) {
    # Mode rapide : on ne sonde que les IP reellement vues dans les tables ARP,
    # filtrees sur le perimetre site/vlan demande.
    $siteSet = @{}; foreach ($s in $SiteList){ $siteSet[$s]=$true }
    $vlanSet = @{}; foreach ($v in $VlanList){ $vlanSet[$v]=$true }
    foreach ($ip in ($MacTable.Keys | Sort-Object)) {
        $o = $ip.Split('.')
        if ($o.Count -ne 4 -or $o[0] -ne "$BaseOctet") { continue }
        if ($siteSet[[int]$o[1]] -and $vlanSet[[int]$o[2]]) { $AllIPs.Add($ip) }
    }
    Write-Host "[MODE] ScanArpOnly : seules les IP moissonnees en ARP sont sondees." -ForegroundColor Cyan
} else {
    if ($ScanArpOnly) { Write-Host "[MODE] ScanArpOnly demande mais aucune table ARP : bascule en balayage complet." -ForegroundColor Yellow }
    foreach ($s in $SiteList) { foreach ($v in $VlanList) { for ($h=$HostStart;$h -le $HostEnd;$h++){ $AllIPs.Add("$BaseOctet.$s.$v.$h") } } }
}
$TotalIPs = $AllIPs.Count
Write-Host ("Total IP a sonder : {0}" -f $TotalIPs) -ForegroundColor Green
if ($TotalIPs -gt 100000) { Write-Host "[!] Volume eleve : pensez a -ScanArpOnly (avec -UseSnmpGateways) ou aux filtres -Sites / -Vlans." -ForegroundColor Yellow }

# --- Worker runspace : ping + ports + banniere + dns (+ SNMP hote optionnel) ---
$WorkerBody = @'
$ping=[Net.NetworkInformation.Ping]::new()
try { $rep=$ping.Send($IPTarget,$PingTimeoutMs); if ($rep.Status -ne 'Success') { return $null } }
catch { return $null } finally { $ping.Dispose() }

$open=New-Object System.Collections.Generic.List[int]
foreach ($pt in $TcpPorts) { if (Test-TcpPort $IPTarget $pt $TcpTimeoutMs) { [void]$open.Add($pt) } }

$srv=""; $title=""; $auth=""
foreach ($sc in @(@{S='http';P=80},@{S='https';P=443},@{S='http';P=8080},@{S='https';P=8443})) {
    if ($open -contains $sc.P) {
        $bn=Get-HttpBanner $sc.S $IPTarget $sc.P $HttpTimeoutSec $PSMajor
        if ($bn) { $srv=$bn.Server; $title=$bn.Title; $auth=$bn.Auth; break }
    }
}

$hostn=""
if ($ResolveDns) {
    try { $iar=[Net.Dns]::BeginGetHostEntry($IPTarget,$null,$null); if ($iar.AsyncWaitHandle.WaitOne(350,$false)) { $he=[Net.Dns]::EndGetHostEntry($iar); $hostn=$he.HostName } } catch {}
}

$sd=""; $sn=""
if ($UseSnmpHosts) {
    $r=Invoke-SnmpRaw $IPTarget $Communaute '1.3.6.1.2.1.1.1.0' $SnmpTimeoutMs 0xA0; if ($r){ $sd=$r.Str }
    $r2=Invoke-SnmpRaw $IPTarget $Communaute '1.3.6.1.2.1.1.5.0' $SnmpTimeoutMs 0xA0; if ($r2){ $sn=$r2.Str }
}

return [PSCustomObject]@{
    IP=$IPTarget; Ports=$open.ToArray(); HttpServer=$srv; HttpTitle=$title; HttpAuth=$auth; Hostname=$(if($hostn){$hostn}elseif($sn){$sn}else{""}); SnmpDescr=$sd
}
'@

$ParamLine = 'param($IPTarget,$PingTimeoutMs,$TcpPorts,$TcpTimeoutMs,$HttpTimeoutSec,$PSMajor,$ResolveDns,$UseSnmpHosts,$Communaute,$SnmpTimeoutMs)'
$ScriptBlock = [scriptblock]::Create($ParamLine + "`n" + $SharedHelpers + "`n" + $SnmpHelpers + "`n" + $WorkerBody)

Write-Host ("Lancement du scan ({0} threads)..." -f $MaxThreads) -ForegroundColor Cyan
$Pool = [runspacefactory]::CreateRunspacePool(1,$MaxThreads); $Pool.Open()
$Common = @{ PingTimeoutMs=$PingTimeoutMs; TcpPorts=$TcpPorts; TcpTimeoutMs=$TcpTimeoutMs; HttpTimeoutSec=$HttpTimeoutSec; PSMajor=$PSMajor; ResolveDns=[bool]$ResolveDns; UseSnmpHosts=[bool]$UseSnmpHosts; Communaute=$Communaute; SnmpTimeoutMs=$SnmpTimeoutMs }

$InFlight = New-Object System.Collections.Generic.List[object]
$Raw      = New-Object System.Collections.Generic.List[object]
$MaxInFlight = $MaxThreads*2; $Index=0; $Done=0

while ($Done -lt $TotalIPs) {
    while ($InFlight.Count -lt $MaxInFlight -and $Index -lt $TotalIPs) {
        $ip=$AllIPs[$Index]; $Index++
        $ps=[powershell]::Create().AddScript($ScriptBlock)
        [void]$ps.AddParameter('IPTarget',$ip); [void]$ps.AddParameters($Common)
        $ps.RunspacePool=$Pool
        $InFlight.Add([PSCustomObject]@{ PS=$ps; Handle=$ps.BeginInvoke() })
    }
    for ($i=$InFlight.Count-1;$i -ge 0;$i--) {
        if ($InFlight[$i].Handle.IsCompleted) {
            try { $res=$InFlight[$i].PS.EndInvoke($InFlight[$i].Handle); if ($res) { foreach($o in $res){ if($o){ $Raw.Add($o) } } } } catch {}
            $InFlight[$i].PS.Dispose(); $InFlight.RemoveAt($i); $Done++
        }
    }
    Write-Progress -Activity "Scan reseau" -Status "$Done / $TotalIPs - $($Raw.Count) hote(s) actif(s)" -PercentComplete ([math]::Round(($Done/$TotalIPs)*100))
    Start-Sleep -Milliseconds 60
}
Write-Progress -Activity "Scan reseau" -Completed
$Pool.Close(); $Pool.Dispose()
Write-Host ("Scan termine. Hotes actifs : {0}" -f $Raw.Count) -ForegroundColor Cyan
if ($Raw.Count -eq 0) { Write-Host "Aucun hote actif detecte." -ForegroundColor Yellow; return }

# --- Enrichissement + classification + conformite ---
Write-Host "Enrichissement (MAC, fabricant, type, conformite)..." -ForegroundColor Cyan
$Inventaire = New-Object System.Collections.Generic.List[object]
foreach ($h in $Raw) {
    $loc = Resolve-Localisation $h.IP
    $mac = ""; if ($MacTable.ContainsKey($h.IP)) { $mac = $MacTable[$h.IP] }
    $vendor = ""; if ($mac) { $oui = Get-Oui $mac; if ($OuiTable.ContainsKey($oui)) { $vendor = $OuiTable[$oui] } }
    $cls  = Get-DeviceType -Vendor $vendor -Ports $h.Ports -HttpServer $h.HttpServer -HttpTitle $h.HttpTitle -HttpAuth $h.HttpAuth -Hostname $h.Hostname -SnmpDescr $h.SnmpDescr
    $conf = Test-Conformite -Type $cls.Type -VlanOctet $loc.VlanOctet -Loc $loc
    $Inventaire.Add([PSCustomObject]@{
        IP           = $h.IP
        Societe      = $loc.Societe
        Ville        = $loc.Ville
        Site         = $loc.SiteId
        VLAN         = $loc.VlanOctet
        VLAN_Usage   = $loc.VlanNom
        MAC          = $mac
        Fabricant    = $vendor
        Type         = $cls.Type
        Confiance    = $cls.Confiance
        Hostname     = $h.Hostname
        Ports        = ($h.Ports -join ',')
        Banniere     = (("{0} {1}" -f $h.HttpServer,$h.HttpTitle).Trim())
        Conformite   = $conf.Statut
        VLAN_Cible   = $conf.VlanCible
        Diagnostic   = $conf.Message
        Signaux      = $cls.Details
    })
}
$Inventaire = $Inventaire | Sort-Object { [version]$_.IP }

# =====================================================================
# ============================= RAPPORTS ==============================
# =====================================================================
Write-Host "`n===================== SYNTHESE =====================" -ForegroundColor Cyan
Write-Host "`n--- Par type d'equipement ---" -ForegroundColor Cyan
$Inventaire | Group-Object Type | Sort-Object Count -Descending | ForEach-Object { Write-Host ("   {0,-14} : {1}" -f $_.Name,$_.Count) }

Write-Host "`n--- Par site ---" -ForegroundColor Cyan
$Inventaire | Group-Object Site | Sort-Object Count -Descending | ForEach-Object { Write-Host ("   {0,-20} : {1}" -f $_.Name,$_.Count) }

Write-Host "`n--- Conformite VLAN ---" -ForegroundColor Cyan
$Inventaire | Group-Object Conformite | Sort-Object Count -Descending | ForEach-Object {
    $c = switch ($_.Name) { 'CONFORME'{'Green'} 'TOLERE'{'Cyan'} 'NON CONFORME'{'Red'} default{'Yellow'} }
    Write-Host ("   {0,-14} : {1}" -f $_.Name,$_.Count) -ForegroundColor $c
}

$Anomalies = @($Inventaire | Where-Object { $_.Conformite -eq 'NON CONFORME' })
if ($Anomalies.Count -gt 0) {
    Write-Host "`n--- NON-CONFORMITES (a corriger) ---" -ForegroundColor Red
    $Anomalies | ForEach-Object { Write-Host ("   {0,-16} {1,-12} {2,-22} -> {3}" -f $_.IP,$_.Type,$_.Fabricant,$_.Diagnostic) -ForegroundColor Red }
}

$SansMac = @($Inventaire | Where-Object { -not $_.MAC })
if ($SansMac.Count -gt 0) {
    Write-Host ("`n[!] {0} hote(s) sans MAC (hors sous-reseau du scanner et non moissonnes en SNMP)." -f $SansMac.Count) -ForegroundColor Yellow
    if (-not $UseSnmpGateways) { Write-Host "    -> relancez avec -UseSnmpGateways pour recuperer les MAC distantes." -ForegroundColor Yellow }
}

# --- Exports ---
$Inventaire | Export-Csv -Path $FichierCSV -NoTypeInformation -Encoding UTF8 -Delimiter ';'
Write-Host ("`n-> CSV  : {0}" -f $FichierCSV) -ForegroundColor Green
$Inventaire | ConvertTo-Json -Depth 4 | Out-File -FilePath $FichierJSON -Encoding UTF8
Write-Host ("-> JSON : {0}" -f $FichierJSON) -ForegroundColor Green
if ($Anomalies.Count -gt 0) {
    $Anomalies | Export-Csv -Path $FichierAnoCSV -NoTypeInformation -Encoding UTF8 -Delimiter ';'
    Write-Host ("-> CSV non-conformites : {0}" -f $FichierAnoCSV) -ForegroundColor Green
}

# --- Excel (si ImportExcel dispo) ---
if (Get-Module -ListAvailable -Name ImportExcel) {
    try {
        if (Test-Path $FichierExcel) { Remove-Item $FichierExcel -Force }
        $total  = $Inventaire.Count
        $nbConf = @($Inventaire | Where-Object { $_.Conformite -eq 'CONFORME' }).Count
        $nbTol  = @($Inventaire | Where-Object { $_.Conformite -eq 'TOLERE' }).Count
        $nbNon  = @($Inventaire | Where-Object { $_.Conformite -eq 'NON CONFORME' }).Count
        $nbNoMac= @($Inventaire | Where-Object { -not $_.MAC }).Count
        $pc = { param($n) if ($total -gt 0) { [math]::Round(($n/$total)*100,1) } else { 0 } }
        $synth = @(
            [PSCustomObject]@{ Indicateur='Total hotes actifs';        Valeur=$total;   'Pourcent'='' }
            [PSCustomObject]@{ Indicateur='Conformes';                 Valeur=$nbConf;  'Pourcent'=(& $pc $nbConf) }
            [PSCustomObject]@{ Indicateur='Toleres';                   Valeur=$nbTol;   'Pourcent'=(& $pc $nbTol) }
            [PSCustomObject]@{ Indicateur='Non conformes';             Valeur=$nbNon;   'Pourcent'=(& $pc $nbNon) }
            [PSCustomObject]@{ Indicateur='Hotes sans MAC';            Valeur=$nbNoMac; 'Pourcent'=(& $pc $nbNoMac) }
            [PSCustomObject]@{ Indicateur='Sites detectes';            Valeur=@($Inventaire | Group-Object Site).Count;  'Pourcent'='' }
            [PSCustomObject]@{ Indicateur='Types distincts';           Valeur=@($Inventaire | Group-Object Type).Count;  'Pourcent'='' }
        )
        $parType = $Inventaire | Group-Object Type | Select-Object @{N='Type';E={$_.Name}},@{N='Nombre';E={$_.Count}} | Sort-Object Nombre -Descending
        $parConf = $Inventaire | Group-Object Conformite | Select-Object @{N='Conformite';E={$_.Name}},@{N='Nombre';E={$_.Count}} | Sort-Object Nombre -Descending
        $parSite = $Inventaire | Group-Object Societe | Select-Object @{N='Site';E={$_.Name}},@{N='Nombre';E={$_.Count}} | Sort-Object Nombre -Descending
        $chart = New-ExcelChartDefinition -Title "Repartition par type" -ChartType BarClustered -XRange "Type" -YRange "Nombre" -NoLegend
        $synth   | Export-Excel -Path $FichierExcel -WorksheetName "KPI" -AutoSize -BoldTopRow -Title ("Inventaire reseau - {0}" -f $Entreprise) -TitleBold -TitleSize 14 -ClearSheet
        $parConf | Export-Excel -Path $FichierExcel -WorksheetName "KPI" -AutoSize -BoldTopRow -StartRow 11
        $parSite | Export-Excel -Path $FichierExcel -WorksheetName "KPI" -AutoSize -BoldTopRow -StartRow 11 -StartColumn 4
        $parType | Export-Excel -Path $FichierExcel -WorksheetName "KPI" -AutoSize -BoldTopRow -StartRow 11 -StartColumn 8 -ExcelChartDefinition $chart
        $xl = $Inventaire | Export-Excel -Path $FichierExcel -WorksheetName "Inventaire" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Medium2 -PassThru
        $ws = $xl.Workbook.Worksheets["Inventaire"]
        $rows = $Inventaire.Count + 1
        # Colonne Conformite = 13 (N) selon l'ordre des proprietes
        Add-ConditionalFormatting -Worksheet $ws -Range ("N2:N$rows") -RuleType ContainsText -ConditionValue "NON CONFORME" -BackgroundColor ([System.Drawing.Color]::FromArgb(255,199,206)) -ForegroundColor ([System.Drawing.Color]::FromArgb(156,0,6))
        Add-ConditionalFormatting -Worksheet $ws -Range ("N2:N$rows") -RuleType ContainsText -ConditionValue "CONFORME"     -BackgroundColor ([System.Drawing.Color]::FromArgb(198,239,206)) -ForegroundColor ([System.Drawing.Color]::FromArgb(0,97,0))
        Add-ConditionalFormatting -Worksheet $ws -Range ("N2:N$rows") -RuleType ContainsText -ConditionValue "TOLERE"       -BackgroundColor ([System.Drawing.Color]::FromArgb(255,235,156)) -ForegroundColor ([System.Drawing.Color]::FromArgb(156,101,0))
        if ($Anomalies.Count -gt 0) { $Anomalies | Export-Excel -Path $FichierExcel -WorksheetName "NonConformites" -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableStyle Light9 }
        Close-ExcelPackage $xl
        Write-Host ("-> Excel: {0}" -f $FichierExcel) -ForegroundColor Green
    } catch { Write-Host ("[Excel] Echec generation : {0}" -f $_) -ForegroundColor Yellow }
} else {
    Write-Host "[INFO] Module ImportExcel absent : export Excel ignore (Install-Module ImportExcel)." -ForegroundColor Yellow
}

Write-Host "`n=== Termine ===" -ForegroundColor Cyan
