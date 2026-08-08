<#
.SYNOPSIS
    Détection complète des objets en double dans Active Directory, avec
    génération d'un rapport Excel (.xlsx) formaté sur le Bureau.

.DESCRIPTION
    Ce script parcourt l'annuaire (utilisateurs, ordinateurs, groupes) et
    détecte plusieurs familles de doublons, utiles à l'hygiène et à la
    sécurité de l'AD :

        1. Objets en conflit (CNF)  : objets « ...CNF:<guid> » créés lors de
           conflits de réplication (vrais objets dupliqués).
        2. SPN en double            : un même ServicePrincipalName porté par
           plusieurs objets (casse l'authentification Kerberos - critique).
        3. UPN en double            : userPrincipalName identiques.
        4. E-mail / proxyAddresses  : adresses de messagerie en collision.
        5. DisplayName en double    : noms d'affichage identiques.
        6. Nom (CN) en double       : même nom d'objet dans des OU différentes.
        7. employeeID en double     : identifiants RH en double.

    Le classeur contient :
        - « Synthèse »  : indicateurs clés (KPI) par famille de doublons.
        - « Doublons »  : le détail, un objet concerné par ligne.

    Toutes les dates sont au format JJ-MM-AAAA.

.PARAMETER SearchBase
    (Optionnel) DN de l'OU de départ. Par défaut : tout le domaine.

.PARAMETER OutputPath
    (Optionnel) Chemin complet du fichier .xlsx. Par défaut : le Bureau.

.EXAMPLE
    .\Export-ADDuplicateObjects.ps1

.NOTES
    Prérequis :
        - Windows avec le module RSAT ActiveDirectory
        - Module ImportExcel (installé automatiquement si absent)
        - Droits de lecture sur l'annuaire Active Directory

    Auteur : Valérian DUFOUR / Claude
#>

[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Doublons_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
)

$ErrorActionPreference = 'Stop'
$Author = 'Valérian DUFOUR / Claude'

function Write-Step { param([string]$Message) Write-Host "[+] $Message" -ForegroundColor Cyan }

# --------------------------------------------------------------------------
# 1. Prérequis
# --------------------------------------------------------------------------
Write-Step "Vérification du module ActiveDirectory..."
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "Le module 'ActiveDirectory' est introuvable. Installez les outils RSAT puis relancez."
}
Import-Module ActiveDirectory -ErrorAction Stop

Write-Step "Vérification du module ImportExcel..."
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Write-Step "Installation du module ImportExcel (utilisateur courant)..."
    try { Install-Module ImportExcel -Scope CurrentUser -Force -AllowClobber -ErrorAction Stop }
    catch { throw "Impossible d'installer ImportExcel. Exécutez : Install-Module ImportExcel -Scope CurrentUser" }
}
Import-Module ImportExcel -ErrorAction Stop

# --------------------------------------------------------------------------
# 2. Fonctions utilitaires
# --------------------------------------------------------------------------
function Get-OUFromDN {
    param([string]$DistinguishedName)
    if ($DistinguishedName -match '^CN=.*?,(.*)$') { return $Matches[1] }
    if ($DistinguishedName -match '^OU=.*?,(.*)$') { return $Matches[1] }
    return $DistinguishedName
}

# --------------------------------------------------------------------------
# 3. Collecte des objets
# --------------------------------------------------------------------------
$baseParam = @{}
if ($SearchBase) { $baseParam['SearchBase'] = $SearchBase; Write-Step "Périmètre : $SearchBase" }

Write-Step "Chargement des utilisateurs..."
$users = Get-ADUser -Filter * -Properties DisplayName, UserPrincipalName, mail, proxyAddresses, `
    employeeID, servicePrincipalName, whenCreated, DistinguishedName, Enabled, SamAccountName @baseParam

Write-Step "Chargement des ordinateurs..."
$computers = Get-ADComputer -Filter * -Properties DisplayName, servicePrincipalName, whenCreated, `
    DistinguishedName, Enabled, SamAccountName @baseParam

Write-Step "Chargement des groupes..."
$groups = Get-ADGroup -Filter * -Properties DisplayName, mail, proxyAddresses, whenCreated, `
    DistinguishedName, SamAccountName @baseParam

# Normalisation en une liste d'objets homogène
function New-Info {
    param($obj, [string]$class)
    [PSCustomObject]@{
        Name        = $obj.Name
        Class       = $class
        Sam         = $obj.SamAccountName
        OU          = (Get-OUFromDN $obj.DistinguishedName)
        Created     = $obj.whenCreated
        Enabled     = if ($null -ne $obj.Enabled) { if ($obj.Enabled) { 'Activé' } else { 'Désactivé' } } else { 'N/A' }
        DN          = $obj.DistinguishedName
        Raw         = $obj
    }
}

$allInfos = [System.Collections.Generic.List[object]]::new()
foreach ($u in $users)     { $allInfos.Add((New-Info $u 'Utilisateur')) }
foreach ($c in $computers) { $allInfos.Add((New-Info $c 'Ordinateur')) }
foreach ($g in $groups)    { $allInfos.Add((New-Info $g 'Groupe')) }
Write-Step ("{0} objet(s) chargé(s) au total." -f $allInfos.Count)

# --------------------------------------------------------------------------
# 4. Détection des doublons
# --------------------------------------------------------------------------
Write-Step "Analyse des doublons..."
$now = Get-Date
$dupRows = [System.Collections.Generic.List[object]]::new()

# Compte, par catégorie : nb de valeurs dupliquées et nb d'objets concernés
$catStats = [ordered]@{}
function Register-Stat { param([string]$Cat, [int]$Values, [int]$Objects)
    $catStats[$Cat] = [PSCustomObject]@{ 'Valeurs en double' = $Values; 'Objets concernés' = $Objects }
}

function Add-DuplicateRows {
    <#
      $Map : hashtable clé (valeur normalisée) -> liste de PSCustomObject info
      Ajoute une ligne par objet pour chaque clé apparaissant > 1 fois.
      Retourne @(nbValeursDupliquees, nbObjetsConcernes)
    #>
    param([string]$TypeLabel, [hashtable]$Map)
    $vals = 0; $objs = 0
    foreach ($key in $Map.Keys) {
        $list = @($Map[$key])
        if ($list.Count -gt 1) {
            $vals++
            foreach ($info in $list) {
                $objs++
                $dupRows.Add([PSCustomObject]@{
                    'Type de doublon'   = $TypeLabel
                    'Valeur dupliquée'  = $key
                    'Nb occurrences'    = $list.Count
                    'Nom de l''objet'   = $info.Name
                    "Type d'objet"      = $info.Class
                    'SamAccountName'    = $info.Sam
                    'Emplacement (OU)'  = $info.OU
                    'Date de création'  = $info.Created
                    'Activé'            = $info.Enabled
                })
            }
        }
    }
    Register-Stat -Cat $TypeLabel -Values $vals -Objects $objs
}

# Fonction générique de construction d'une map (clé normalisée -> infos)
function Build-Map {
    param([scriptblock]$Selector)   # renvoie une ou plusieurs clés pour un info
    $map = @{}
    foreach ($info in $allInfos) {
        foreach ($k in (& $Selector $info)) {
            if ([string]::IsNullOrWhiteSpace($k)) { continue }
            $kk = "$k".Trim().ToLower()
            if (-not $map.ContainsKey($kk)) { $map[$kk] = [System.Collections.Generic.List[object]]::new() }
            $map[$kk].Add($info)
        }
    }
    return $map
}

# 4.2 SPN en double (sécurité - Kerberos)
$spnMap = Build-Map { param($i) @($i.Raw.servicePrincipalName) }
Add-DuplicateRows -TypeLabel 'SPN en double' -Map $spnMap

# 4.3 UPN en double
$upnMap = Build-Map { param($i) ,$i.Raw.UserPrincipalName }
Add-DuplicateRows -TypeLabel 'UPN en double' -Map $upnMap

# 4.4 E-mail / proxyAddresses en double
$mailMap = Build-Map {
    param($i)
    $keys = @()
    if ($i.Raw.mail) { $keys += $i.Raw.mail }
    foreach ($p in @($i.Raw.proxyAddresses)) {
        if ($p) { $keys += ($p -replace '^(?i)smtp:', '') }
    }
    $keys
}
Add-DuplicateRows -TypeLabel 'E-mail en double' -Map $mailMap

# 4.5 DisplayName en double
$dispMap = Build-Map { param($i) ,$i.Raw.DisplayName }
Add-DuplicateRows -TypeLabel 'DisplayName en double' -Map $dispMap

# 4.6 Nom (CN) en double (dans des OU différentes)
$nameMap = Build-Map { param($i) ,$i.Name }
Add-DuplicateRows -TypeLabel 'Nom (CN) en double' -Map $nameMap

# 4.7 employeeID en double
$empMap = Build-Map { param($i) ,$i.Raw.employeeID }
Add-DuplicateRows -TypeLabel 'employeeID en double' -Map $empMap

# 4.1 Objets en conflit de réplication (CNF)
$cnfCount = 0
try {
    $cnfParam = @{ LDAPFilter = '(|(cn=*CNF:*)(ou=*CNF:*)(name=*CNF:*))'
                   Properties = @('whenCreated','objectClass') }
    if ($SearchBase) { $cnfParam['SearchBase'] = $SearchBase }
    $cnfObjects = Get-ADObject @cnfParam
    foreach ($o in @($cnfObjects)) {
        $cnfCount++
        $dupRows.Add([PSCustomObject]@{
            'Type de doublon'   = 'Objet en conflit (CNF)'
            'Valeur dupliquée'  = $o.Name
            'Nb occurrences'    = 1
            'Nom de l''objet'   = $o.Name
            "Type d'objet"      = "$($o.objectClass)"
            'SamAccountName'    = ''
            'Emplacement (OU)'  = (Get-OUFromDN $o.DistinguishedName)
            'Date de création'  = $o.whenCreated
            'Activé'            = 'N/A'
        })
    }
} catch { Write-Warning "Recherche des objets CNF impossible : $($_.Exception.Message)" }
Register-Stat -Cat 'Objet en conflit (CNF)' -Values $cnfCount -Objects $cnfCount

$dupReport = $dupRows | Sort-Object 'Type de doublon','Valeur dupliquée','Nom de l''objet'
Write-Step ("{0} ligne(s) de doublon détectée(s)." -f @($dupReport).Count)

# --------------------------------------------------------------------------
# 5. Export Excel
# --------------------------------------------------------------------------
Write-Step "Génération du fichier Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

# S'il n'y a aucun doublon, on écrit tout de même une ligne "propre"
if (@($dupReport).Count -eq 0) {
    $dupReport = ,([PSCustomObject]@{
        'Type de doublon'  = 'Aucun doublon détecté'
        'Valeur dupliquée' = ''
        'Nb occurrences'   = 0
        'Nom de l''objet'  = ''
        "Type d'objet"     = ''
        'SamAccountName'   = ''
        'Emplacement (OU)' = ''
        'Date de création' = $null
        'Activé'           = ''
    })
}

$excel = $dupReport | Export-Excel -Path $OutputPath -WorksheetName 'Doublons' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'Doublons' -TableStyle 'Medium2' `
    -Title ("Doublons Active Directory - {0}" -f $now.ToString('dd-MM-yyyy HH:mm')) -TitleBold -TitleSize 14 -PassThru

$wsD = $excel.Workbook.Worksheets['Doublons']
$hrD = $wsD.Dimension.Start.Row
# Colonne H = Date de création
for ($c = 1; $c -le $wsD.Dimension.End.Column; $c++) {
    if ($wsD.Cells[$hrD, $c].Value -eq 'Date de création') {
        $l = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c)
        $wsD.Cells[("{0}{1}:{0}{2}" -f $l, ($hrD + 1), $wsD.Dimension.End.Row)].Style.Numberformat.Format = 'dd-mm-yyyy hh:mm'
    }
}
$rangeD = "A{0}:{1}{2}" -f ($hrD + 1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($wsD.Dimension.End.Column), $wsD.Dimension.End.Row
# Type = colonne A : SPN (critique) en rouge, CNF en orange
Add-ConditionalFormatting -Worksheet $wsD -Range $rangeD -RuleType Expression `
    -ConditionValue ('=$A{0}="SPN en double"' -f ($hrD + 1)) -BackgroundColor ([System.Drawing.Color]::MistyRose) -Bold
Add-ConditionalFormatting -Worksheet $wsD -Range $rangeD -RuleType Expression `
    -ConditionValue ('=$A{0}="Objet en conflit (CNF)"' -f ($hrD + 1)) -BackgroundColor ([System.Drawing.Color]::Moccasin)

# --- Feuille Synthèse ---
$totalObjectsConcerned = (@($catStats.Values | Measure-Object 'Objets concernés' -Sum).Sum)
if ($null -eq $totalObjectsConcerned) { $totalObjectsConcerned = 0 }

$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse' -MoveToStart
$wsK.Cells['B2'].Value = 'Rapport des objets en double - Active Directory'
$wsK.Cells['B2'].Style.Font.Size = 16
$wsK.Cells['B2'].Style.Font.Bold = $true
$wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f $now.ToString('dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$wsK.Cells['B5'].Value = ("Objets analysés : {0}" -f $allInfos.Count)

# En-tête tableau KPI
$row = 7
$wsK.Cells["B$row"].Value = 'Famille de doublon'
$wsK.Cells["C$row"].Value = 'Valeurs en double'
$wsK.Cells["D$row"].Value = 'Objets concernés'
$wsK.Cells["B$row:D$row"].Style.Font.Bold = $true
$wsK.Cells["B$row:D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$row:D$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

foreach ($cat in $catStats.Keys) {
    $row++
    $wsK.Cells["B$row"].Value = $cat
    $wsK.Cells["C$row"].Value = $catStats[$cat].'Valeurs en double'
    $wsK.Cells["D$row"].Value = $catStats[$cat].'Objets concernés'
    if ($row % 2 -eq 0) {
        $wsK.Cells["B$row:D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217,226,243))
    }
    # SPN en rouge (critique)
    if ($cat -eq 'SPN en double' -and $catStats[$cat].'Objets concernés' -gt 0) {
        $wsK.Cells["B$row:D$row"].Style.Font.Color.SetColor([System.Drawing.Color]::DarkRed)
        $wsK.Cells["B$row:D$row"].Style.Font.Bold = $true
    }
}
$firstCatRow = 8
$lastCatRow  = $row

# Ligne total
$row++
$wsK.Cells["B$row"].Value = 'TOTAL objets concernés'
$wsK.Cells["D$row"].Value = $totalObjectsConcerned
$wsK.Cells["B$row:D$row"].Style.Font.Bold = $true

try {
    $chart = $wsK.Drawings.AddChart('dupChart', [OfficeOpenXml.Drawing.Chart.eChartType]::BarClustered)
    $chart.Title.Text = 'Objets concernés par famille de doublon'
    $chart.SetPosition(6, 0, 5, 0); $chart.SetSize(520, 300)
    $serie = $chart.Series.Add($wsK.Cells["D$firstCatRow:D$lastCatRow"], $wsK.Cells["B$firstCatRow:B$lastCatRow"])
    $serie.Header = 'Objets concernés'
    $chart.DataLabel.ShowValue = $true
} catch { Write-Warning "Graphique non généré : $($_.Exception.Message)" }

$wsK.Column(2).Width = 32; $wsK.Column(3).Width = 18; $wsK.Column(4).Width = 18

Close-ExcelPackage $excel

Write-Host ""
Write-Host "[OK] Rapport généré : $OutputPath" -ForegroundColor Green
Write-Host ("     {0} objet(s) concerné(s) par un doublon." -f $totalObjectsConcerned) -ForegroundColor Green
