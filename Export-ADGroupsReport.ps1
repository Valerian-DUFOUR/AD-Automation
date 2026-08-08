<#
.SYNOPSIS
    Extraction des groupes Active Directory et de leurs membres, avec
    génération d'un rapport Excel (.xlsx) formaté sur le Bureau.

.DESCRIPTION
    Ce script interroge Active Directory via le module ActiveDirectory
    (Get-ADGroup) et produit un classeur Excel à plusieurs feuilles :

        1. « Synthèse »            : indicateurs clés (KPI) de l'extraction.
        2. « Groupes »            : un groupe par ligne, avec sa description,
           ses notes (« ce que fait le groupe »), sa catégorie, son étendue,
           sa date de création, son emplacement (OU) et le nombre de membres
           ventilé par type (utilisateurs / ordinateurs / groupes).
        3. « Membres Utilisateurs » : un couple (groupe, utilisateur) par
           ligne, avec les mêmes champs que le script d'extraction des
           utilisateurs.
        4. « Membres Ordinateurs »  : un couple (groupe, ordinateur) par
           ligne, avec les mêmes champs que le script d'extraction des
           ordinateurs (dont l'obsolescence Windows).

    Toutes les dates sont au format JJ-MM-AAAA.

.PARAMETER SearchBase
    (Optionnel) DN de l'OU de départ pour limiter la recherche des groupes.
    Exemple : "OU=Groupes,DC=example,DC=local"
    Par défaut : tout le domaine.

.PARAMETER InactiveDays
    (Optionnel) Seuil, en jours, d'inactivité d'un ordinateur. Défaut : 90.

.PARAMETER OutputPath
    (Optionnel) Chemin complet du fichier .xlsx de sortie.
    Par défaut : le Bureau de l'utilisateur courant.

.EXAMPLE
    .\Export-ADGroupsReport.ps1

.EXAMPLE
    .\Export-ADGroupsReport.ps1 -SearchBase "OU=Groupes,DC=example,DC=local"

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

    [int]$InactiveDays = 90,

    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Groupes_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
)

$ErrorActionPreference = 'Stop'
$Author = 'Valérian DUFOUR / Claude'

function Write-Step { param([string]$Message) Write-Host "[+] $Message" -ForegroundColor Cyan }

# --------------------------------------------------------------------------
# 1. Prérequis
# --------------------------------------------------------------------------
Write-Step "Vérification du module ActiveDirectory..."
if (-not (Get-Module -ListAvailable -Name ActiveDirectory)) {
    throw "Le module 'ActiveDirectory' est introuvable. Installez les outils RSAT puis relancez le script."
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
    return $DistinguishedName
}

function Get-ObsolescenceStatus {
    param([string]$OperatingSystem)
    if ([string]::IsNullOrWhiteSpace($OperatingSystem)) {
        return [PSCustomObject]@{ Obsolete = 'Inconnu'; Motif = 'OS non renseigné' }
    }
    $os = $OperatingSystem
    if ($os -match 'Windows 11')    { return [PSCustomObject]@{ Obsolete='Non'; Motif='Support en cours' } }
    if ($os -match 'Windows 10')    { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support (14/10/2025)' } }
    if ($os -match 'Windows 8')     { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support (Windows 8/8.1)' } }
    if ($os -match 'Windows 7')     { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support (14/01/2020)' } }
    if ($os -match 'Windows Vista') { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support' } }
    if ($os -match 'Windows XP')    { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support' } }
    if ($os -match 'Server 2025')   { return [PSCustomObject]@{ Obsolete='Non'; Motif='Support en cours' } }
    if ($os -match 'Server 2022')   { return [PSCustomObject]@{ Obsolete='Non'; Motif='Support en cours' } }
    if ($os -match 'Server 2019')   { return [PSCustomObject]@{ Obsolete='Non'; Motif='Support en cours' } }
    if ($os -match 'Server 2016')   { return [PSCustomObject]@{ Obsolete='Non'; Motif='Support étendu (jusqu''au 12/01/2027)' } }
    if ($os -match 'Server 2012')   { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support (10/10/2023)' } }
    if ($os -match 'Server 2008')   { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support (14/01/2020)' } }
    if ($os -match 'Server 2003')   { return [PSCustomObject]@{ Obsolete='Oui'; Motif='Fin de support' } }
    return [PSCustomObject]@{ Obsolete='À vérifier'; Motif='OS non répertorié' }
}

function Translate-Category { param($c) switch ("$c") { 'Security' {'Sécurité'} 'Distribution' {'Distribution'} default {"$c"} } }
function Translate-Scope    { param($s) switch ("$s") { 'DomainLocal' {'Domaine local'} 'Global' {'Global'} 'Universal' {'Universel'} default {"$s"} } }

# --------------------------------------------------------------------------
# 3. Pré-chargement des utilisateurs et ordinateurs (cache par DN)
# --------------------------------------------------------------------------
Write-Step "Chargement des utilisateurs (cache)..."
$userProps = @('SamAccountName','mail','GivenName','Surname','DisplayName','Description',
               'LastLogonDate','PasswordLastSet','whenCreated','DistinguishedName',
               'Enabled','PasswordNeverExpires')
$userByDN = @{}
Get-ADUser -Filter * -Properties $userProps | ForEach-Object { $userByDN[$_.DistinguishedName] = $_ }
Write-Step ("{0} utilisateur(s) en cache." -f $userByDN.Count)

Write-Step "Chargement des ordinateurs (cache)..."
$compProps = @('Name','Description','OperatingSystem','OperatingSystemVersion',
               'whenCreated','LastLogonDate','Enabled','DistinguishedName')
$compByDN = @{}
Get-ADComputer -Filter * -Properties $compProps | ForEach-Object { $compByDN[$_.DistinguishedName] = $_ }
Write-Step ("{0} ordinateur(s) en cache." -f $compByDN.Count)

# --------------------------------------------------------------------------
# 4. Récupération des groupes
# --------------------------------------------------------------------------
Write-Step "Interrogation des groupes Active Directory..."
$groupProps = @('Description','Info','whenCreated','GroupCategory','GroupScope',
                'member','DistinguishedName','ManagedBy','mail')
$getParams = @{ Filter = '*'; Properties = $groupProps }
if ($SearchBase) { $getParams['SearchBase'] = $SearchBase; Write-Step "Périmètre : $SearchBase" }

$groups = Get-ADGroup @getParams
$groupDNset = [System.Collections.Generic.HashSet[string]]::new()
$groups | ForEach-Object { [void]$groupDNset.Add($_.DistinguishedName) }
Write-Step ("{0} groupe(s) récupéré(s)." -f $groups.Count)

# --------------------------------------------------------------------------
# 5. Construction des jeux de données
# --------------------------------------------------------------------------
Write-Step "Analyse des membres et préparation du rapport..."
$now = Get-Date

$groupRows = [System.Collections.Generic.List[object]]::new()
$userRows  = [System.Collections.Generic.List[object]]::new()
$compRows  = [System.Collections.Generic.List[object]]::new()

foreach ($g in $groups) {
    $members = @($g.member | Where-Object { $_ })
    $nbUser = 0; $nbComp = 0; $nbGroup = 0; $nbOther = 0

    foreach ($dn in $members) {
        if ($userByDN.ContainsKey($dn)) {
            $nbUser++
            $u = $userByDN[$dn]
            $userRows.Add([PSCustomObject]@{
                'Groupe'                 = $g.Name
                'Nom du compte'          = $u.SamAccountName
                'Email'                  = $u.mail
                'Prénom'                 = $u.GivenName
                'Nom'                    = $u.Surname
                'Nom complet'            = $u.DisplayName
                'Description'            = $u.Description
                'Dernière connexion'     = $u.LastLogonDate
                'Dernier changement MDP' = $u.PasswordLastSet
                'Date de création'       = $u.whenCreated
                'Emplacement (OU)'       = (Get-OUFromDN $u.DistinguishedName)
                'Statut'                 = if ($u.Enabled) { 'Activé' } else { 'Désactivé' }
                'MDP n''expire jamais'   = if ($u.PasswordNeverExpires) { 'Oui' } else { 'Non' }
            })
        }
        elseif ($compByDN.ContainsKey($dn)) {
            $nbComp++
            $c = $compByDN[$dn]
            $daysInactive = if ($c.LastLogonDate) { [int]($now - $c.LastLogonDate).TotalDays } else { $null }
            $obs = Get-ObsolescenceStatus -OperatingSystem $c.OperatingSystem
            $compRows.Add([PSCustomObject]@{
                'Groupe'             = $g.Name
                'Nom du PC'          = $c.Name
                'Description'        = $c.Description
                'OS'                 = $c.OperatingSystem
                'Version OS'         = $c.OperatingSystemVersion
                'Date de création'   = $c.whenCreated
                'Dernière connexion' = $c.LastLogonDate
                'Statut'             = if ($c.Enabled) { 'Activé' } else { 'Désactivé' }
                'Jours inactif'      = $daysInactive
                'Emplacement (OU)'   = (Get-OUFromDN $c.DistinguishedName)
                'Obsolète'           = $obs.Obsolete
                'Motif obsolescence' = $obs.Motif
            })
        }
        elseif ($groupDNset.Contains($dn)) { $nbGroup++ }
        else { $nbOther++ }
    }

    $groupRows.Add([PSCustomObject]@{
        'Nom du groupe'      = $g.Name
        'Description'        = $g.Description
        'Notes (rôle)'       = $g.Info
        'Catégorie'          = (Translate-Category $g.GroupCategory)
        'Étendue'            = (Translate-Scope $g.GroupScope)
        'Géré par'           = if ($g.ManagedBy) { $g.ManagedBy -replace '^CN=([^,]+),.*$','$1' } else { '' }
        'Email'              = $g.mail
        'Date de création'   = $g.whenCreated
        'Nb membres'         = $members.Count
        'Utilisateurs'       = $nbUser
        'Ordinateurs'        = $nbComp
        'Groupes'            = $nbGroup
        'Autres'             = $nbOther
        'Emplacement (OU)'   = (Get-OUFromDN $g.DistinguishedName)
    })
}

$groupReport = $groupRows | Sort-Object 'Nom du groupe'
$userReport  = $userRows  | Sort-Object 'Groupe','Nom du compte'
$compReport  = $compRows  | Sort-Object 'Groupe','Nom du PC'

# --------------------------------------------------------------------------
# 6. KPI
# --------------------------------------------------------------------------
$totalGroups = @($groupReport).Count
$secGroups   = @($groupReport | Where-Object { $_.'Catégorie' -eq 'Sécurité' }).Count
$distGroups  = @($groupReport | Where-Object { $_.'Catégorie' -eq 'Distribution' }).Count
$emptyGroups = @($groupReport | Where-Object { $_.'Nb membres' -eq 0 }).Count
$totalMembers= (@($groupReport | Measure-Object 'Nb membres' -Sum).Sum)
if ($null -eq $totalMembers) { $totalMembers = 0 }
$totalUsers  = @($userReport).Count
$totalComps  = @($compReport).Count

$pct = { param($n) if ($totalGroups -gt 0) { [math]::Round(($n / $totalGroups) * 100, 1) } else { 0 } }

$scopeDist = $groupReport | Group-Object 'Étendue' | Sort-Object Count -Descending |
    ForEach-Object { [PSCustomObject]@{ 'Étendue' = $_.Name; 'Nombre' = $_.Count } }

# --------------------------------------------------------------------------
# 7. Export Excel
# --------------------------------------------------------------------------
Write-Step "Génération du fichier Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

$dateFmt = 'dd-mm-yyyy hh:mm'

# --- Feuille Groupes ---
$excel = $groupReport | Export-Excel -Path $OutputPath -WorksheetName 'Groupes' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'Groupes' -TableStyle 'Medium2' `
    -Title ("Groupes Active Directory - {0}" -f $now.ToString('dd-MM-yyyy HH:mm')) -TitleBold -TitleSize 14 -PassThru

function Format-DateColumns {
    param($ws, [string[]]$Columns)
    $hr = $ws.Dimension.Start.Row
    for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
        if ($Columns -contains $ws.Cells[$hr, $c].Value) {
            $l = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c)
            $ws.Cells[("{0}{1}:{0}{2}" -f $l, ($hr + 1), $ws.Dimension.End.Row)].Style.Numberformat.Format = $dateFmt
        }
    }
}

$wsG = $excel.Workbook.Worksheets['Groupes']
Format-DateColumns -ws $wsG -Columns @('Date de création')
$hrG = $wsG.Dimension.Start.Row
$rangeG = "A{0}:{1}{2}" -f ($hrG + 1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($wsG.Dimension.End.Column), $wsG.Dimension.End.Row
# Groupes vides mis en évidence (colonne I = Nb membres)
Add-ConditionalFormatting -Worksheet $wsG -Range $rangeG -RuleType Expression `
    -ConditionValue ('=$I{0}=0' -f ($hrG + 1)) -BackgroundColor ([System.Drawing.Color]::MistyRose)

# --- Feuille Membres Utilisateurs ---
if (@($userReport).Count -gt 0) {
    $excel = $userReport | Export-Excel -ExcelPackage $excel -WorksheetName 'Membres Utilisateurs' `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'MembresUtilisateurs' -TableStyle 'Medium2' -PassThru
    $wsU = $excel.Workbook.Worksheets['Membres Utilisateurs']
    Format-DateColumns -ws $wsU -Columns @('Dernière connexion','Dernier changement MDP','Date de création')
    $hrU = $wsU.Dimension.Start.Row
    $rangeU = "A{0}:{1}{2}" -f ($hrU + 1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($wsU.Dimension.End.Column), $wsU.Dimension.End.Row
    # Colonne L = Statut, M = MDP n'expire jamais
    Add-ConditionalFormatting -Worksheet $wsU -Range $rangeU -RuleType Expression `
        -ConditionValue ('=$L{0}="Désactivé"' -f ($hrU + 1)) -ForegroundColor ([System.Drawing.Color]::Gray) -Italic
    Add-ConditionalFormatting -Worksheet $wsU -Range $rangeU -RuleType Expression `
        -ConditionValue ('=$M{0}="Oui"' -f ($hrU + 1)) -ForegroundColor ([System.Drawing.Color]::DarkRed) -Bold
}

# --- Feuille Membres Ordinateurs ---
if (@($compReport).Count -gt 0) {
    $excel = $compReport | Export-Excel -ExcelPackage $excel -WorksheetName 'Membres Ordinateurs' `
        -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'MembresOrdinateurs' -TableStyle 'Medium2' -PassThru
    $wsC = $excel.Workbook.Worksheets['Membres Ordinateurs']
    Format-DateColumns -ws $wsC -Columns @('Date de création','Dernière connexion')
    $hrC = $wsC.Dimension.Start.Row
    $rangeC = "A{0}:{1}{2}" -f ($hrC + 1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($wsC.Dimension.End.Column), $wsC.Dimension.End.Row
    # Colonne K = Obsolète, H = Statut
    Add-ConditionalFormatting -Worksheet $wsC -Range $rangeC -RuleType Expression `
        -ConditionValue ('=$K{0}="Oui"' -f ($hrC + 1)) -BackgroundColor ([System.Drawing.Color]::MistyRose)
    Add-ConditionalFormatting -Worksheet $wsC -Range $rangeC -RuleType Expression `
        -ConditionValue ('=$H{0}="Désactivé"' -f ($hrC + 1)) -ForegroundColor ([System.Drawing.Color]::Gray) -Italic
}

# --- Feuille Synthèse (1ère position) ---
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse' -MoveToStart

$wsK.Cells['B2'].Value = 'Rapport des groupes Active Directory'
$wsK.Cells['B2'].Style.Font.Size = 16
$wsK.Cells['B2'].Style.Font.Bold = $true
$wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f $now.ToString('dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)

$kpis = @(
    @('Total des groupes',            $totalGroups,  $null),
    @('Groupes de sécurité',          $secGroups,    (& $pct $secGroups)),
    @('Groupes de distribution',      $distGroups,   (& $pct $distGroups)),
    @('Groupes vides (0 membre)',     $emptyGroups,  (& $pct $emptyGroups)),
    @('Total des appartenances',      $totalMembers, $null),
    @('  dont utilisateurs',          $totalUsers,   $null),
    @('  dont ordinateurs',           $totalComps,   $null)
)

$row = 6
$wsK.Cells["B$row"].Value = 'Indicateur'; $wsK.Cells["C$row"].Value = 'Valeur'; $wsK.Cells["D$row"].Value = '%'
$wsK.Cells["B$row:D$row"].Style.Font.Bold = $true
$wsK.Cells["B$row:D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$row:D$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
foreach ($k in $kpis) {
    $row++
    $wsK.Cells["B$row"].Value = $k[0]; $wsK.Cells["C$row"].Value = $k[1]
    if ($null -ne $k[2]) { $wsK.Cells["D$row"].Value = ("{0} %" -f $k[2]) }
    if ($row % 2 -eq 0) {
        $wsK.Cells["B$row:D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217,226,243))
    }
}
$kpiLastRow = $row

# Répartition par étendue
$scopeHeaderRow = $kpiLastRow + 3
$wsK.Cells["B$scopeHeaderRow"].Value = 'Étendue'; $wsK.Cells["C$scopeHeaderRow"].Value = 'Nombre'
$wsK.Cells["B$scopeHeaderRow:C$scopeHeaderRow"].Style.Font.Bold = $true
$wsK.Cells["B$scopeHeaderRow:C$scopeHeaderRow"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$scopeHeaderRow:C$scopeHeaderRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$scopeHeaderRow:C$scopeHeaderRow"].Style.Font.Color.SetColor([System.Drawing.Color]::White)
$r = $scopeHeaderRow
foreach ($s in $scopeDist) { $r++; $wsK.Cells["B$r"].Value = $s.'Étendue'; $wsK.Cells["C$r"].Value = $s.'Nombre' }
$scopeLastRow = $r

try {
    $chart = $wsK.Drawings.AddChart('scopeChart', [OfficeOpenXml.Drawing.Chart.eChartType]::BarClustered)
    $chart.Title.Text = 'Répartition des groupes par étendue'
    $chart.SetPosition(5, 0, 5, 0); $chart.SetSize(460, 280)
    $serie = $chart.Series.Add($wsK.Cells["C$($scopeHeaderRow + 1):C$scopeLastRow"], $wsK.Cells["B$($scopeHeaderRow + 1):B$scopeLastRow"])
    $serie.Header = 'Nombre de groupes'
    $chart.DataLabel.ShowValue = $true
} catch { Write-Warning "Graphique non généré : $($_.Exception.Message)" }

$wsK.Column(2).Width = 34; $wsK.Column(3).Width = 14; $wsK.Column(4).Width = 10

Close-ExcelPackage $excel

Write-Host ""
Write-Host "[OK] Rapport généré : $OutputPath" -ForegroundColor Green
Write-Host ("     {0} groupe(s) | {1} appartenance(s) utilisateur | {2} appartenance(s) ordinateur." -f `
    $totalGroups, $totalUsers, $totalComps) -ForegroundColor Green
