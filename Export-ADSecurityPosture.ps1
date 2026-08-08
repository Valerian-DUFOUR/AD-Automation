<#
.SYNOPSIS
    Évaluation de la posture de sécurité d'Active Directory (type PingCastle
    « allégé ») avec génération d'un rapport Excel (.xlsx) sur le Bureau.

.DESCRIPTION
    Ce script produit un tableau de bord orienté RSSI :

        - Feuille « Synthèse (KPI) » : les indicateurs clés avec un niveau de
          risque (OK / À surveiller / Critique) et un graphique de répartition.
        - Une feuille de DÉTAIL par thème, pour faciliter la lecture :
            * Comptes privilégiés   (membres des groupes d'administration)
            * Comptes inactifs      (utilisateurs activés sans connexion récente)
            * Ordinateurs inactifs
            * Problèmes de mots de passe (n'expire jamais, non requis, ancien)
            * Comptes Kerberoastables (utilisateurs porteurs d'un SPN)
            * Délégation             (non contrainte / contrainte)
            * SIDHistory            (objets avec un historique de SID)
            * Politique de mots de passe (domaine + stratégies affinées)

    Toutes les dates sont au format JJ-MM-AAAA. Le script est en LECTURE SEULE :
    il n'effectue aucune modification dans l'annuaire.

.PARAMETER InactiveDays
    Seuil d'inactivité (jours). Défaut : 90.

.PARAMETER MaxPasswordAgeDays
    Âge de mot de passe considéré comme trop ancien (jours). Défaut : 365.

.PARAMETER OutputPath
    Chemin complet du .xlsx. Par défaut : le Bureau.

.EXAMPLE
    .\Export-ADSecurityPosture.ps1

.NOTES
    Prérequis :
        - Windows avec le module RSAT ActiveDirectory
        - Module ImportExcel (installé automatiquement si absent)
        - Droits de lecture sur l'annuaire Active Directory

    Avertissement : les seuils et pondérations sont indicatifs et doivent être
    adaptés à votre contexte. Ce script ne remplace pas un audit complet
    (PingCastle, PurpleKnight, etc.).

    Auteur : Valérian DUFOUR / Claude
#>

[CmdletBinding()]
param(
    [int]$InactiveDays = 90,
    [int]$MaxPasswordAgeDays = 365,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_Securite_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
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
function Days-Since {
    param($Date)
    if ($Date) { return [int]((Get-Date) - $Date).TotalDays } else { return $null }
}

# --------------------------------------------------------------------------
# 3. Collecte
# --------------------------------------------------------------------------
$now       = Get-Date
$domain    = Get-ADDomain
$domainSID = $domain.DomainSID.Value
Write-Step ("Domaine : {0}" -f $domain.DNSRoot)

Write-Step "Chargement des utilisateurs..."
$uProps = @('DisplayName','SamAccountName','Enabled','LastLogonDate','PasswordLastSet',
            'PasswordNeverExpires','PasswordNotRequired','whenCreated','DistinguishedName',
            'servicePrincipalName','TrustedForDelegation','TrustedToAuthForDelegation',
            'msDS-AllowedToDelegateTo','SIDHistory','adminCount')
$users = Get-ADUser -Filter * -Properties $uProps
$userByDN = @{}
$users | ForEach-Object { $userByDN[$_.DistinguishedName] = $_ }
Write-Step ("{0} utilisateur(s)." -f $users.Count)

Write-Step "Chargement des ordinateurs..."
$cProps = @('Name','SamAccountName','Enabled','LastLogonDate','OperatingSystem','whenCreated',
            'DistinguishedName','TrustedForDelegation','TrustedToAuthForDelegation',
            'msDS-AllowedToDelegateTo','SIDHistory')
$computers = Get-ADComputer -Filter * -Properties $cProps
Write-Step ("{0} ordinateur(s)." -f $computers.Count)

# --------------------------------------------------------------------------
# 4. Comptes privilégiés
# --------------------------------------------------------------------------
Write-Step "Analyse des groupes privilégiés..."
$privGroups = @(
    @{ Name='Domain Admins';                SID="$domainSID-512" },
    @{ Name='Enterprise Admins';            SID="$domainSID-519" },
    @{ Name='Schema Admins';                SID="$domainSID-518" },
    @{ Name='Group Policy Creator Owners';  SID="$domainSID-520" },
    @{ Name='Administrators';               SID='S-1-5-32-544' },
    @{ Name='Account Operators';            SID='S-1-5-32-548' },
    @{ Name='Backup Operators';             SID='S-1-5-32-551' },
    @{ Name='Server Operators';             SID='S-1-5-32-549' },
    @{ Name='Print Operators';              SID='S-1-5-32-550' }
)

$privRows    = [System.Collections.Generic.List[object]]::new()
$privUserDNs = [System.Collections.Generic.HashSet[string]]::new()
$privCount   = @{}

foreach ($pg in $privGroups) {
    $grp = $null
    try { $grp = Get-ADGroup -Identity $pg.SID -ErrorAction Stop } catch { continue }
    $members = @()
    try { $members = @(Get-ADGroupMember -Identity $grp -Recursive -ErrorAction SilentlyContinue) } catch {}
    $userMembers = @($members | Where-Object { $_.objectClass -eq 'user' })
    $privCount[$pg.Name] = $userMembers.Count
    foreach ($m in $userMembers) {
        [void]$privUserDNs.Add($m.distinguishedName)
        $u = $userByDN[$m.distinguishedName]
        if ($u) {
            $privRows.Add([PSCustomObject]@{
                'Groupe privilégié'      = $pg.Name
                'Nom du compte'          = $u.SamAccountName
                'Nom complet'            = $u.DisplayName
                'Activé'                 = if ($u.Enabled) { 'Activé' } else { 'Désactivé' }
                'Dernière connexion'     = $u.LastLogonDate
                'Jours inactif'          = (Days-Since $u.LastLogonDate)
                'MDP n''expire jamais'   = if ($u.PasswordNeverExpires) { 'Oui' } else { 'Non' }
                'Dernier changement MDP' = $u.PasswordLastSet
                'Emplacement (OU)'       = (Get-OUFromDN $u.DistinguishedName)
            })
        } else {
            $privRows.Add([PSCustomObject]@{
                'Groupe privilégié'      = $pg.Name
                'Nom du compte'          = $m.SamAccountName
                'Nom complet'            = $m.name
                'Activé'                 = 'N/A (externe)'
                'Dernière connexion'     = $null
                'Jours inactif'          = $null
                'MDP n''expire jamais'   = 'N/A'
                'Dernier changement MDP' = $null
                'Emplacement (OU)'       = (Get-OUFromDN $m.distinguishedName)
            })
        }
    }
}
$privReport = $privRows | Sort-Object 'Groupe privilégié','Nom du compte'
$privTotal  = $privUserDNs.Count
$privNeverExp = @($privReport | Where-Object { $_.'MDP n''expire jamais' -eq 'Oui' }).Count
$privInactive = @($privReport | Where-Object { $null -ne $_.'Jours inactif' -and $_.'Jours inactif' -ge $InactiveDays }).Count

# --------------------------------------------------------------------------
# 5. Comptes / ordinateurs inactifs
# --------------------------------------------------------------------------
Write-Step "Analyse des objets inactifs..."
$staleUsers = $users | Where-Object {
    $_.Enabled -and (
        ($_.LastLogonDate -and ((Days-Since $_.LastLogonDate) -ge $InactiveDays)) -or
        (-not $_.LastLogonDate -and $_.whenCreated -and ((Days-Since $_.whenCreated) -ge $InactiveDays))
    )
} | ForEach-Object {
    [PSCustomObject]@{
        'Nom du compte'      = $_.SamAccountName
        'Nom complet'        = $_.DisplayName
        'Dernière connexion' = $_.LastLogonDate
        'Jours inactif'      = (Days-Since $_.LastLogonDate)
        'Date de création'   = $_.whenCreated
        'Emplacement (OU)'   = (Get-OUFromDN $_.DistinguishedName)
    }
} | Sort-Object 'Jours inactif' -Descending

$staleComputers = $computers | Where-Object {
    $_.Enabled -and (
        ($_.LastLogonDate -and ((Days-Since $_.LastLogonDate) -ge $InactiveDays)) -or
        (-not $_.LastLogonDate -and $_.whenCreated -and ((Days-Since $_.whenCreated) -ge $InactiveDays))
    )
} | ForEach-Object {
    [PSCustomObject]@{
        'Nom du PC'          = $_.Name
        'OS'                 = $_.OperatingSystem
        'Dernière connexion' = $_.LastLogonDate
        'Jours inactif'      = (Days-Since $_.LastLogonDate)
        'Date de création'   = $_.whenCreated
        'Emplacement (OU)'   = (Get-OUFromDN $_.DistinguishedName)
    }
} | Sort-Object 'Jours inactif' -Descending

# --------------------------------------------------------------------------
# 6. Problèmes de mots de passe
# --------------------------------------------------------------------------
Write-Step "Analyse des mots de passe..."
$pwdIssues = $users | Where-Object {
    $_.PasswordNeverExpires -or $_.PasswordNotRequired -or
    (-not $_.PasswordLastSet) -or
    ($_.PasswordLastSet -and ((Days-Since $_.PasswordLastSet) -ge $MaxPasswordAgeDays))
} | ForEach-Object {
    [PSCustomObject]@{
        'Nom du compte'         = $_.SamAccountName
        'Nom complet'           = $_.DisplayName
        'Activé'                = if ($_.Enabled) { 'Activé' } else { 'Désactivé' }
        'MDP n''expire jamais'  = if ($_.PasswordNeverExpires) { 'Oui' } else { 'Non' }
        'MDP non requis'        = if ($_.PasswordNotRequired) { 'Oui' } else { 'Non' }
        'MDP jamais défini'     = if (-not $_.PasswordLastSet) { 'Oui' } else { 'Non' }
        'Âge MDP (jours)'       = (Days-Since $_.PasswordLastSet)
        'Dernier changement'    = $_.PasswordLastSet
        'Emplacement (OU)'      = (Get-OUFromDN $_.DistinguishedName)
    }
} | Sort-Object 'Âge MDP (jours)' -Descending

$pwdNeverExp = @($users | Where-Object { $_.PasswordNeverExpires -and $_.Enabled }).Count
$pwdNotReq   = @($users | Where-Object { $_.PasswordNotRequired -and $_.Enabled }).Count

# --------------------------------------------------------------------------
# 7. Comptes Kerberoastables (SPN sur des comptes utilisateurs)
# --------------------------------------------------------------------------
Write-Step "Analyse des comptes Kerberoastables..."
$kerberoast = $users | Where-Object {
    $_.SamAccountName -ne 'krbtgt' -and @($_.servicePrincipalName).Count -gt 0
} | ForEach-Object {
    [PSCustomObject]@{
        'Nom du compte'        = $_.SamAccountName
        'Nom complet'          = $_.DisplayName
        'Activé'               = if ($_.Enabled) { 'Activé' } else { 'Désactivé' }
        'MDP n''expire jamais' = if ($_.PasswordNeverExpires) { 'Oui' } else { 'Non' }
        'Âge MDP (jours)'      = (Days-Since $_.PasswordLastSet)
        'SPN'                  = (@($_.servicePrincipalName) -join ' ; ')
        'Emplacement (OU)'     = (Get-OUFromDN $_.DistinguishedName)
    }
} | Sort-Object 'Nom du compte'

# --------------------------------------------------------------------------
# 8. Délégation
# --------------------------------------------------------------------------
Write-Step "Analyse des délégations..."
$delegation = [System.Collections.Generic.List[object]]::new()
foreach ($o in @($users) + @($computers)) {
    $type = $null
    if ($o.TrustedForDelegation) { $type = 'Non contrainte (CRITIQUE)' }
    elseif ($o.TrustedToAuthForDelegation) { $type = 'Contrainte avec transition de protocole' }
    elseif (@($o.'msDS-AllowedToDelegateTo').Count -gt 0) { $type = 'Contrainte' }
    if ($type) {
        $delegation.Add([PSCustomObject]@{
            'Objet'            = $o.SamAccountName
            "Type d'objet"     = if ($o.objectClass -eq 'computer') { 'Ordinateur' } else { 'Utilisateur' }
            'Type délégation'  = $type
            'Cibles'           = (@($o.'msDS-AllowedToDelegateTo') -join ' ; ')
            'Activé'           = if ($o.Enabled) { 'Activé' } else { 'Désactivé' }
            'Emplacement (OU)' = (Get-OUFromDN $o.DistinguishedName)
        })
    }
}
$delegReport = $delegation | Sort-Object 'Type délégation','Objet'
$unconstrained = @($delegReport | Where-Object { $_.'Type délégation' -like 'Non contrainte*' }).Count

# --------------------------------------------------------------------------
# 9. SIDHistory
# --------------------------------------------------------------------------
Write-Step "Analyse des SIDHistory..."
$sidHistory = [System.Collections.Generic.List[object]]::new()
foreach ($o in @($users) + @($computers)) {
    if (@($o.SIDHistory).Count -gt 0) {
        $sidHistory.Add([PSCustomObject]@{
            'Objet'            = $o.SamAccountName
            "Type d'objet"     = if ($o.objectClass -eq 'computer') { 'Ordinateur' } else { 'Utilisateur' }
            'Nb SIDHistory'    = @($o.SIDHistory).Count
            'SIDHistory'       = (@($o.SIDHistory) -join ' ; ')
            'Emplacement (OU)' = (Get-OUFromDN $o.DistinguishedName)
        })
    }
}
$sidReport = $sidHistory | Sort-Object 'Objet'

# --------------------------------------------------------------------------
# 10. Politique de mots de passe
# --------------------------------------------------------------------------
Write-Step "Lecture des politiques de mot de passe..."
$pwdPolicyRows = [System.Collections.Generic.List[object]]::new()
try {
    $dp = Get-ADDefaultDomainPasswordPolicy
    $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'='Domaine (par défaut)'; 'Paramètre'='Longueur minimale';        'Valeur'=$dp.MinPasswordLength })
    $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'='Domaine (par défaut)'; 'Paramètre'='Complexité activée';       'Valeur'=$dp.ComplexityEnabled })
    $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'='Domaine (par défaut)'; 'Paramètre'='Âge max (jours)';          'Valeur'=$dp.MaxPasswordAge.Days })
    $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'='Domaine (par défaut)'; 'Paramètre'='Âge min (jours)';          'Valeur'=$dp.MinPasswordAge.Days })
    $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'='Domaine (par défaut)'; 'Paramètre'='Historique conservé';      'Valeur'=$dp.PasswordHistoryCount })
    $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'='Domaine (par défaut)'; 'Paramètre'='Seuil verrouillage';       'Valeur'=$dp.LockoutThreshold })
    $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'='Domaine (par défaut)'; 'Paramètre'='Durée verrouillage (min)'; 'Valeur'=$dp.LockoutDuration.TotalMinutes })
} catch { Write-Warning "Politique de domaine illisible : $($_.Exception.Message)" }
try {
    foreach ($fg in @(Get-ADFineGrainedPasswordPolicy -Filter * -ErrorAction SilentlyContinue)) {
        $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'=$fg.Name; 'Paramètre'='Longueur minimale';  'Valeur'=$fg.MinPasswordLength })
        $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'=$fg.Name; 'Paramètre'='Complexité activée'; 'Valeur'=$fg.ComplexityEnabled })
        $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'=$fg.Name; 'Paramètre'='Âge max (jours)';    'Valeur'=$fg.MaxPasswordAge.Days })
        $pwdPolicyRows.Add([PSCustomObject]@{ 'Politique'=$fg.Name; 'Paramètre'='Priorité';           'Valeur'=$fg.Precedence })
    }
} catch {}

# --------------------------------------------------------------------------
# 11. Indicateurs additionnels (krbtgt, invité)
# --------------------------------------------------------------------------
$krbtgt = $users | Where-Object { $_.SamAccountName -eq 'krbtgt' } | Select-Object -First 1
$krbtgtAge = if ($krbtgt) { (Days-Since $krbtgt.PasswordLastSet) } else { $null }

$guest = $users | Where-Object { $_.SamAccountName -eq 'Guest' } | Select-Object -First 1
$guestEnabled = if ($guest -and $guest.Enabled) { 'Oui' } else { 'Non' }

# --------------------------------------------------------------------------
# 12. Export Excel : feuilles de détail
# --------------------------------------------------------------------------
Write-Step "Génération du fichier Excel : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }

$script:excel = $null
function Add-Sheet {
    param($Data, [string]$SheetName, [string]$TableName, [string[]]$DateColumns)
    if (@($Data).Count -eq 0) { $Data = ,([PSCustomObject]@{ 'Information' = 'Aucun élément détecté' }) }
    if ($null -eq $script:excel) {
        $script:excel = $Data | Export-Excel -Path $OutputPath -WorksheetName $SheetName `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $TableName -TableStyle 'Medium2' -PassThru
    } else {
        $script:excel = $Data | Export-Excel -ExcelPackage $script:excel -WorksheetName $SheetName `
            -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $TableName -TableStyle 'Medium2' -PassThru
    }
    if ($DateColumns) {
        $ws = $script:excel.Workbook.Worksheets[$SheetName]
        $hr = $ws.Dimension.Start.Row
        for ($c = 1; $c -le $ws.Dimension.End.Column; $c++) {
            if ($DateColumns -contains $ws.Cells[$hr, $c].Value) {
                $l = [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c)
                $ws.Cells[("{0}{1}:{0}{2}" -f $l, ($hr + 1), $ws.Dimension.End.Row)].Style.Numberformat.Format = 'dd-mm-yyyy hh:mm'
            }
        }
    }
}

Add-Sheet -Data $privReport     -SheetName 'Comptes privilégiés'  -TableName 'CptsPrivilegies'  -DateColumns @('Dernière connexion','Dernier changement MDP')
Add-Sheet -Data $staleUsers     -SheetName 'Comptes inactifs'     -TableName 'CptsInactifs'     -DateColumns @('Dernière connexion','Date de création')
Add-Sheet -Data $staleComputers -SheetName 'Ordinateurs inactifs' -TableName 'PCInactifs'       -DateColumns @('Dernière connexion','Date de création')
Add-Sheet -Data $pwdIssues      -SheetName 'Problèmes MDP'        -TableName 'ProblemesMDP'     -DateColumns @('Dernier changement')
Add-Sheet -Data $kerberoast     -SheetName 'Kerberoastables'      -TableName 'Kerberoastables'
Add-Sheet -Data $delegReport    -SheetName 'Délégation'           -TableName 'Delegation'
Add-Sheet -Data $sidReport      -SheetName 'SIDHistory'           -TableName 'SIDHistory'
Add-Sheet -Data $pwdPolicyRows  -SheetName 'Politique MDP'        -TableName 'PolitiqueMDP'
$excel = $script:excel

# Surlignages critiques : délégation non contrainte
$wsDel = $excel.Workbook.Worksheets['Délégation']
if (@($delegReport).Count -gt 0 -and $wsDel.Dimension) {
    $hr = $wsDel.Dimension.Start.Row
    $rg = "A{0}:{1}{2}" -f ($hr + 1), [OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($wsDel.Dimension.End.Column), $wsDel.Dimension.End.Row
    # Colonne C = Type délégation
    Add-ConditionalFormatting -Worksheet $wsDel -Range $rg -RuleType Expression `
        -ConditionValue ('=ISNUMBER(SEARCH("Non contrainte",$C{0}))' -f ($hr + 1)) -BackgroundColor ([System.Drawing.Color]::MistyRose) -Bold
}

# --------------------------------------------------------------------------
# 13. Feuille Synthèse (KPI) en 1ère position
# --------------------------------------------------------------------------
Write-Step "Construction de la synthèse KPI..."

function Level-Threshold {
    param([double]$Value, [double]$Warn, [double]$Crit)
    if ($Value -ge $Crit) { 'Critique' } elseif ($Value -ge $Warn) { 'À surveiller' } else { 'OK' }
}

$kpi = @(
    @('Membres Domain Admins',              ($privCount['Domain Admins']),      (Level-Threshold ([double]($privCount['Domain Admins']))      5 10)),
    @('Membres Enterprise Admins',          ($privCount['Enterprise Admins']),  (Level-Threshold ([double]($privCount['Enterprise Admins']))  1 3)),
    @('Membres Schema Admins',              ($privCount['Schema Admins']),      (Level-Threshold ([double]($privCount['Schema Admins']))      1 2)),
    @('Total comptes privilégiés (uniques)',$privTotal,                          (Level-Threshold ([double]$privTotal)                         20 50)),
    @('Comptes privilégiés MDP éternel',    $privNeverExp,                       (Level-Threshold ([double]$privNeverExp)                      1 3)),
    @('Comptes privilégiés inactifs',       $privInactive,                       (Level-Threshold ([double]$privInactive)                      1 3)),
    @("Âge mot de passe krbtgt (jours)",    $krbtgtAge,                          (Level-Threshold ([double]([int]$krbtgtAge))                  180 365)),
    @("Comptes inactifs (> $InactiveDays j)",@($staleUsers).Count,               (Level-Threshold ([double](@($staleUsers).Count))            1 50)),
    @("Ordinateurs inactifs (> $InactiveDays j)",@($staleComputers).Count,       (Level-Threshold ([double](@($staleComputers).Count))        1 50)),
    @("MDP n'expire jamais (activés)",      $pwdNeverExp,                        (Level-Threshold ([double]$pwdNeverExp)                       1 20)),
    @('MDP non requis (activés)',           $pwdNotReq,                          (Level-Threshold ([double]$pwdNotReq)                         1 1)),
    @('Comptes Kerberoastables (SPN)',      @($kerberoast).Count,                (Level-Threshold ([double](@($kerberoast).Count))            1 5)),
    @('Délégation non contrainte',          $unconstrained,                      (Level-Threshold ([double]$unconstrained)                    1 1)),
    @('Objets avec SIDHistory',             @($sidReport).Count,                 (Level-Threshold ([double](@($sidReport).Count))             1 10)),
    @('Compte invité (Guest) activé',       $guestEnabled,                       (if ($guestEnabled -eq 'Oui') { 'Critique' } else { 'OK' }))
)

$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Synthèse (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Posture de sécurité Active Directory'
$wsK.Cells['B2'].Style.Font.Size = 16
$wsK.Cells['B2'].Style.Font.Bold = $true
$wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Domaine : {0}" -f $domain.DNSRoot)
$wsK.Cells['B4'].Value = ("Date d'extraction : {0}" -f $now.ToString('dd-MM-yyyy HH:mm'))
$wsK.Cells['B5'].Value = ("Auteur : {0}" -f $Author)

$hdrRow = 7
$wsK.Cells["B$hdrRow"].Value = 'Indicateur'
$wsK.Cells["C$hdrRow"].Value = 'Valeur'
$wsK.Cells["D$hdrRow"].Value = 'Niveau'
$wsK.Cells["B$hdrRow:D$hdrRow"].Style.Font.Bold = $true
$wsK.Cells["B$hdrRow:D$hdrRow"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
$wsK.Cells["B$hdrRow:D$hdrRow"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
$wsK.Cells["B$hdrRow:D$hdrRow"].Style.Font.Color.SetColor([System.Drawing.Color]::White)

$colOK   = [System.Drawing.Color]::FromArgb(198,239,206)
$colWarn = [System.Drawing.Color]::FromArgb(255,235,156)
$colCrit = [System.Drawing.Color]::FromArgb(255,199,206)
$nOK = 0; $nWarn = 0; $nCrit = 0

$row = $hdrRow
foreach ($k in $kpi) {
    $row++
    $wsK.Cells["B$row"].Value = $k[0]
    $wsK.Cells["C$row"].Value = $k[1]
    $wsK.Cells["D$row"].Value = $k[2]
    $col = switch ($k[2]) { 'Critique' { $colCrit; $nCrit++ } 'À surveiller' { $colWarn; $nWarn++ } default { $colOK; $nOK++ } }
    $wsK.Cells["D$row"].Style.Fill.PatternType = [OfficeOpenXml.Style.ExcelFillStyle]::Solid
    $wsK.Cells["D$row"].Style.Fill.BackgroundColor.SetColor($col)
    $wsK.Cells["D$row"].Style.Font.Bold = $true
}

# Petit tableau + graphique de répartition des niveaux
$sumRow = $row + 3
$wsK.Cells["B$sumRow"].Value = 'Niveau'; $wsK.Cells["C$sumRow"].Value = 'Nombre'
$wsK.Cells["B$sumRow:C$sumRow"].Style.Font.Bold = $true
$wsK.Cells["B$($sumRow+1)"].Value = 'OK';           $wsK.Cells["C$($sumRow+1)"].Value = $nOK
$wsK.Cells["B$($sumRow+2)"].Value = 'À surveiller'; $wsK.Cells["C$($sumRow+2)"].Value = $nWarn
$wsK.Cells["B$($sumRow+3)"].Value = 'Critique';     $wsK.Cells["C$($sumRow+3)"].Value = $nCrit

try {
    $chart = $wsK.Drawings.AddChart('riskChart', [OfficeOpenXml.Drawing.Chart.eChartType]::Pie)
    $chart.Title.Text = 'Répartition des indicateurs par niveau'
    $chart.SetPosition(6, 0, 5, 0); $chart.SetSize(420, 300)
    $null = $chart.Series.Add($wsK.Cells["C$($sumRow+1):C$($sumRow+3)"], $wsK.Cells["B$($sumRow+1):B$($sumRow+3)"])
    $chart.DataLabel.ShowValue = $true
} catch { Write-Warning "Graphique non généré : $($_.Exception.Message)" }

$wsK.Column(2).Width = 40; $wsK.Column(3).Width = 16; $wsK.Column(4).Width = 16

Close-ExcelPackage $excel

Write-Host ""
Write-Host "[OK] Rapport généré : $OutputPath" -ForegroundColor Green
Write-Host ("     Indicateurs -> OK: {0} | À surveiller: {1} | Critique: {2}" -f $nOK, $nWarn, $nCrit) -ForegroundColor Green
