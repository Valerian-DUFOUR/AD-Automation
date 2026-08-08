<#
.SYNOPSIS
    Rapport des comptes utilisateurs et des ordinateurs inactifs dans Active
    Directory, avec une page KPI par tranche d'inactivité, exporté en Excel
    (.xlsx) sur le Bureau.

.DESCRIPTION
    Reprend les informations du rapport des utilisateurs et du rapport des
    ordinateurs, mais filtré sur les objets INACTIFS (> 30 jours), et ajoute :

        - Feuille « Synthèse (KPI) » : nombre de comptes et d'ordinateurs
          inactifs de plus de 30, 90, 180 et 365 jours (cumulatif), avec un
          graphique comparatif.
        - Feuille « Utilisateurs inactifs » : mêmes colonnes que le rapport
          des utilisateurs + jours d'inactivité + tranche.
        - Feuille « Ordinateurs inactifs » : mêmes colonnes que le rapport des
          ordinateurs + tranche.

    L'inactivité est mesurée sur la dernière connexion (LastLogonDate). Si
    celle-ci est vide, la date de création du compte est utilisée comme repère
    (pour ne pas signaler à tort un objet récent jamais connecté).

    Toutes les dates sont au format JJ-MM-AAAA.

.PARAMETER SearchBase
    (Optionnel) DN de l'OU de départ. Par défaut : tout le domaine.

.PARAMETER OutputPath
    (Optionnel) Chemin complet du .xlsx. Par défaut : le Bureau.

.EXAMPLE
    .\Export-ADInactiveReport.ps1

.NOTES
    Prérequis : RSAT ActiveDirectory + module ImportExcel (auto-installé).
    Auteur : Valérian DUFOUR / Claude
#>

[CmdletBinding()]
param(
    [string]$SearchBase,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Inactifs_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
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
    return $DistinguishedName
}
function Get-Tranche {
    param([int]$Days)
    if ($Days -ge 365) { '365 jours et +' }
    elseif ($Days -ge 180) { '180-364 jours' }
    elseif ($Days -ge 90)  { '90-179 jours' }
    elseif ($Days -ge 30)  { '30-89 jours' }
    else { '< 30 jours' }
}
function Get-ObsolescenceStatus {
    param([string]$OperatingSystem)
    if ([string]::IsNullOrWhiteSpace($OperatingSystem)) { return [PSCustomObject]@{ Obsolete='Inconnu'; Motif='OS non renseigné' } }
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

# --------------------------------------------------------------------------
# 3. Collecte
# --------------------------------------------------------------------------
$now = Get-Date
$baseParam = @{}
if ($SearchBase) { $baseParam['SearchBase'] = $SearchBase; Write-Step "Périmètre : $SearchBase" }

Write-Step "Chargement des utilisateurs..."
$users = Get-ADUser -Filter * -Properties SamAccountName, mail, GivenName, Surname, DisplayName, `
    Description, LastLogonDate, PasswordLastSet, whenCreated, DistinguishedName, Enabled, PasswordNeverExpires @baseParam

Write-Step "Chargement des ordinateurs..."
$computers = Get-ADComputer -Filter * -Properties Name, Description, OperatingSystem, OperatingSystemVersion, `
    whenCreated, LastLogonDate, DistinguishedName, Enabled @baseParam

# --------------------------------------------------------------------------
# 4. Préparation - utilisateurs inactifs (> 30 j)
# --------------------------------------------------------------------------
Write-Step "Analyse de l'inactivité..."
$userInactive = foreach ($u in $users) {
    $ref = if ($u.LastLogonDate) { $u.LastLogonDate } else { $u.whenCreated }
    if (-not $ref) { continue }
    $days = [int]($now - $ref).TotalDays
    if ($days -lt 30) { continue }
    [PSCustomObject]@{
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
        'Jours inactif'          = $days
        'Tranche'                = (Get-Tranche $days)
    }
}
$userInactive = @($userInactive) | Sort-Object 'Jours inactif' -Descending

# --------------------------------------------------------------------------
# 5. Préparation - ordinateurs inactifs (> 30 j)
# --------------------------------------------------------------------------
$compInactive = foreach ($c in $computers) {
    $ref = if ($c.LastLogonDate) { $c.LastLogonDate } else { $c.whenCreated }
    if (-not $ref) { continue }
    $days = [int]($now - $ref).TotalDays
    if ($days -lt 30) { continue }
    $obs = Get-ObsolescenceStatus -OperatingSystem $c.OperatingSystem
    [PSCustomObject]@{
        'Nom du PC'          = $c.Name
        'Description'        = $c.Description
        'OS'                 = $c.OperatingSystem
        'Version OS'         = $c.OperatingSystemVersion
        'Date de création'   = $c.whenCreated
        'Dernière connexion' = $c.LastLogonDate
        'Statut'             = if ($c.Enabled) { 'Activé' } else { 'Désactivé' }
        'Jours inactif'      = $days
        'Emplacement (OU)'   = (Get-OUFromDN $c.DistinguishedName)
        'Obsolète'           = $obs.Obsolete
        'Motif obsolescence' = $obs.Motif
        'Tranche'            = (Get-Tranche $days)
    }
}
$compInactive = @($compInactive) | Sort-Object 'Jours inactif' -Descending

Write-Step ("{0} utilisateur(s) et {1} ordinateur(s) inactifs (> 30 j)." -f $userInactive.Count, $compInactive.Count)

# --------------------------------------------------------------------------
# 6. KPI cumulatifs
# --------------------------------------------------------------------------
$seuils = @(30, 90, 180, 365)
function Count-AtLeast { param($Data, [int]$Seuil) @($Data | Where-Object { $_.'Jours inactif' -ge $Seuil }).Count }

# --------------------------------------------------------------------------
# 7. Export Excel
# --------------------------------------------------------------------------
Write-Step "Génération du fichier Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$dateFmt = 'dd-mm-yyyy hh:mm'

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

$dataU = if ($userInactive.Count -gt 0) { $userInactive } else { ,([PSCustomObject]@{ 'Information' = 'Aucun utilisateur inactif' }) }
$excel = $dataU | Export-Excel -Path $OutputPath -WorksheetName 'Utilisateurs inactifs' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'UtilisateursInactifs' -TableStyle 'Medium2' -PassThru
if ($userInactive.Count -gt 0) {
    $wsU = $excel.Workbook.Worksheets['Utilisateurs inactifs']
    Format-DateColumns -ws $wsU -Columns @('Dernière connexion','Dernier changement MDP','Date de création')
}

$dataC = if ($compInactive.Count -gt 0) { $compInactive } else { ,([PSCustomObject]@{ 'Information' = 'Aucun ordinateur inactif' }) }
$excel = $dataC | Export-Excel -ExcelPackage $excel -WorksheetName 'Ordinateurs inactifs' `
    -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName 'OrdinateursInactifs' -TableStyle 'Medium2' -PassThru
if ($compInactive.Count -gt 0) {
    $wsC = $excel.Workbook.Worksheets['Ordinateurs inactifs']
    Format-DateColumns -ws $wsC -Columns @('Date de création','Dernière connexion')
}

# --- Feuille Synthèse (KPI) ---
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Comptes et ordinateurs inactifs - Active Directory'
$wsK.Cells['B2'].Style.Font.Size = 16
$wsK.Cells['B2'].Style.Font.Bold = $true
$wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Date d'extraction : {0}" -f $now.ToString('dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$wsK.Cells['B5'].Value = "Comptages cumulatifs (un objet inactif > 365 j est aussi compté dans > 30 / 90 / 180 j)."

$hdr = 7
$wsK.Cells["B$hdr"].Value = "Seuil d'inactivité"
$wsK.Cells["C$hdr"].Value = 'Utilisateurs'
$wsK.Cells["D$hdr"].Value = 'Ordinateurs'
$wsK.Cells["B$hdr:D$hdr"].Style.Font.Bold = $true
$wsK.Cells["B$hdr:D$hdr"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$hdr:D$hdr"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$hdr:D$hdr"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

$row = $hdr
foreach ($s in $seuils) {
    $row++
    $wsK.Cells["B$row"].Value = ("Plus de {0} jours" -f $s)
    $wsK.Cells["C$row"].Value = (Count-AtLeast $userInactive $s)
    $wsK.Cells["D$row"].Value = (Count-AtLeast $compInactive $s)
    if ($row % 2 -eq 0) {
        $wsK.Cells["B$row:D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(217,226,243))
    }
}
$firstRow = $hdr + 1
$lastRow  = $row

# Total ligne
$row++
$wsK.Cells["B$row"].Value = 'Total inactifs (> 30 j)'
$wsK.Cells["C$row"].Value = $userInactive.Count
$wsK.Cells["D$row"].Value = $compInactive.Count
$wsK.Cells["B$row:D$row"].Style.Font.Bold = $true

try {
    $chart = $wsK.Drawings.AddChart('inactChart', [OfficeOpenXml.Drawing.Chart.eChartType]::ColumnClustered)
    $chart.Title.Text = "Objets inactifs par seuil (cumulatif)"
    $chart.SetPosition(6, 0, 5, 0); $chart.SetSize(560, 320)
    $s1 = $chart.Series.Add($wsK.Cells["C$firstRow:C$lastRow"], $wsK.Cells["B$firstRow:B$lastRow"]); $s1.Header = 'Utilisateurs'
    $s2 = $chart.Series.Add($wsK.Cells["D$firstRow:D$lastRow"], $wsK.Cells["B$firstRow:B$lastRow"]); $s2.Header = 'Ordinateurs'
    $chart.DataLabel.ShowValue = $true
} catch { Write-Warning "Graphique non généré : $($_.Exception.Message)" }

$wsK.Column(2).Width = 24; $wsK.Column(3).Width = 16; $wsK.Column(4).Width = 16

Close-ExcelPackage $excel

Write-Host ""
Write-Host "[OK] Rapport généré : $OutputPath" -ForegroundColor Green
Write-Host ("     Utilisateurs inactifs: {0} | Ordinateurs inactifs: {1}" -f $userInactive.Count, $compInactive.Count) -ForegroundColor Green
