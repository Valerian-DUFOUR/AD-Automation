<#
.SYNOPSIS
    Rapport Active Directory TOUT-EN-UN (version complète) : un seul classeur
    Excel (.xlsx) avec un tableau de bord KPI exécutif et ~25 feuilles de détail
    couvrant en profondeur l'ensemble de l'annuaire.

.DESCRIPTION
    Script auto-suffisant qui collecte l'AD en une passe et produit un « giga
    rapport » unique. Feuilles :
        Tableau de bord (KPI), Utilisateurs, Ordinateurs, Comptes inactifs,
        Ordinateurs inactifs, Comptes privilégiés, Kerberoastables, MDP à risque,
        Délégation (comptes), Délégations ACL (DCSync), SIDHistory, Groupes,
        Membres Utilisateurs, Membres Ordinateurs, Trusts, GPO, Liens GPO,
        cpassword (GPP), LAPS, Sans BitLocker, Doublons, OU/Groupes vides,
        Échéances, FSMO, Santé DC, Réplication, AD CS.

    Modules optionnels absents (GroupPolicy, ADCS…) ignorés sans erreur. Dates au
    format JJ-MM-AAAA. LECTURE SEULE.

    Hors périmètre (autres contextes d'exécution — voir scripts dédiés) : scan
    SNMP, inventaire réseau, audit SMB/NTFS, protocoles vulnérables, capture live
    LDAP, comparaison temporelle.

.PARAMETER SearchBase
    OU de départ pour comptes/ordinateurs/groupes (défaut : tout le domaine).
.PARAMETER InactiveDays
    Seuil d'inactivité (défaut : 90).
.PARAMETER MaxPasswordAgeDays
    Âge de mot de passe jugé trop ancien (défaut : 365).
.PARAMETER SkipGroupMembers
    N'inclut pas le détail des membres par groupe (utile sur très gros domaines).
.PARAMETER SkipAclScan
    N'exécute pas l'analyse des délégations ACL (plus lente).
.PARAMETER OutputPath
    Chemin du .xlsx (défaut : Bureau).

.EXAMPLE
    .\Export-ADFullReport.ps1
    .\Export-ADFullReport.ps1 -InactiveDays 60 -SkipGroupMembers

.NOTES
    Prérequis : RSAT ActiveDirectory (+ GroupPolicy pour les GPO) + ImportExcel
    (auto-installé). À lancer sur un DC avec des droits d'administration.
    Auteur : Valérian DUFOUR / Claude
#>
[CmdletBinding()]
param(
    [string]$SearchBase,
    [int]$InactiveDays = 90,
    [int]$MaxPasswordAgeDays = 365,
    [switch]$SkipGroupMembers,
    [switch]$SkipAclScan,
    [string]$OutputPath = (Join-Path ([Environment]::GetFolderPath('Desktop')) `
        ("Rapport_AD_COMPLET_{0}.xlsx" -f (Get-Date -Format 'dd-MM-yyyy_HHmm')))
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

function Get-OUFromDN { param($d) if ($d -match '^CN=.*?,(.*)$') { return $Matches[1] }; if ($d -match '^OU=.*?,(.*)$') { return $Matches[1] }; return $d }
function Days-Since { param($Date) if ($Date) { return [int]((Get-Date) - $Date).TotalDays } else { return $null } }
function Get-ObsolescenceStatus {
    param([string]$os)
    if ([string]::IsNullOrWhiteSpace($os)) { return 'Inconnu' }
    if ($os -match 'Windows 11') { return 'Non' }
    if ($os -match 'Windows 10|Windows 8|Windows 7|Windows Vista|Windows XP|Server 2012|Server 2008|Server 2003') { return 'Oui' }
    if ($os -match 'Server 2025|Server 2022|Server 2019|Server 2016') { return 'Non' }
    return 'À vérifier'
}

$now = Get-Date
$domain = Get-ADDomain
$forest = Get-ADForest
$baseParam = @{}; if ($SearchBase) { $baseParam['SearchBase'] = $SearchBase }

# ==========================================================================
# COLLECTE
# ==========================================================================
Write-Step "Chargement des utilisateurs..."
$uProps = @('SamAccountName','mail','GivenName','Surname','DisplayName','Description','LastLogonDate','PasswordLastSet',
            'whenCreated','DistinguishedName','Enabled','PasswordNeverExpires','PasswordNotRequired','servicePrincipalName',
            'adminCount','UserPrincipalName','proxyAddresses','AccountExpirationDate','msDS-UserPasswordExpiryTimeComputed',
            'SIDHistory','TrustedForDelegation','TrustedToAuthForDelegation','msDS-AllowedToDelegateTo','employeeID','Name')
$users = Get-ADUser -Filter * -Properties $uProps @baseParam
$userByDN = @{}; $users | ForEach-Object { $userByDN[$_.DistinguishedName] = $_ }

Write-Step "Chargement des ordinateurs..."
$cProps = @('Name','Description','OperatingSystem','OperatingSystemVersion','whenCreated','LastLogonDate','DistinguishedName',
            'Enabled','SIDHistory','TrustedForDelegation','TrustedToAuthForDelegation','msDS-AllowedToDelegateTo')
$computers = Get-ADComputer -Filter * -Properties $cProps @baseParam
$compByDN = @{}; $computers | ForEach-Object { $compByDN[$_.DistinguishedName] = $_ }

Write-Step "Chargement des groupes..."
$groups = Get-ADGroup -Filter * -Properties GroupCategory,GroupScope,member,whenCreated,Description,DistinguishedName,Info,ManagedBy,mail @baseParam
$groupDNset = New-Object System.Collections.Generic.HashSet[string]; $groups | ForEach-Object { [void]$groupDNset.Add($_.DistinguishedName) }

# ---- Utilisateurs / Ordinateurs (détail) ----
$usersReport = $users | ForEach-Object {
    [PSCustomObject]@{ 'Nom du compte'=$_.SamAccountName; 'Email'=$_.mail; 'Prénom'=$_.GivenName; 'Nom'=$_.Surname; 'Nom complet'=$_.DisplayName
        'Description'=$_.Description; 'Dernière connexion'=$_.LastLogonDate; 'Dernier changement MDP'=$_.PasswordLastSet
        'Date de création'=$_.whenCreated; 'Emplacement (OU)'=(Get-OUFromDN $_.DistinguishedName)
        'Statut'=(if($_.Enabled){'Activé'}else{'Désactivé'}); "MDP n'expire jamais"=(if($_.PasswordNeverExpires){'Oui'}else{'Non'}) }
} | Sort-Object 'Nom du compte'
$compReport = $computers | ForEach-Object {
    [PSCustomObject]@{ 'Nom du PC'=$_.Name; 'Description'=$_.Description; 'OS'=$_.OperatingSystem; 'Version OS'=$_.OperatingSystemVersion
        'Date de création'=$_.whenCreated; 'Dernière connexion'=$_.LastLogonDate; 'Statut'=(if($_.Enabled){'Activé'}else{'Désactivé'})
        'Jours inactif'=(Days-Since $_.LastLogonDate); 'Emplacement (OU)'=(Get-OUFromDN $_.DistinguishedName); 'Obsolète'=(Get-ObsolescenceStatus $_.OperatingSystem) }
} | Sort-Object 'Nom du PC'

# ---- Inactifs ----
$userInactive = foreach ($u in $users) { if (-not $u.Enabled) { continue }; $ref = if($u.LastLogonDate){$u.LastLogonDate}else{$u.whenCreated}; $d=Days-Since $ref; if ($d -ge $InactiveDays) { [PSCustomObject]@{ 'Nom du compte'=$u.SamAccountName; 'Nom complet'=$u.DisplayName; 'Dernière connexion'=$u.LastLogonDate; 'Jours inactif'=$d; 'Date de création'=$u.whenCreated; 'Emplacement (OU)'=(Get-OUFromDN $u.DistinguishedName) } } }
$userInactive = @($userInactive) | Sort-Object 'Jours inactif' -Descending
$compInactive = foreach ($c in $computers) { if (-not $c.Enabled) { continue }; $ref = if($c.LastLogonDate){$c.LastLogonDate}else{$c.whenCreated}; $d=Days-Since $ref; if ($d -ge $InactiveDays) { [PSCustomObject]@{ 'Nom du PC'=$c.Name; 'OS'=$c.OperatingSystem; 'Dernière connexion'=$c.LastLogonDate; 'Jours inactif'=$d; 'Emplacement (OU)'=(Get-OUFromDN $c.DistinguishedName); 'Obsolète'=(Get-ObsolescenceStatus $c.OperatingSystem) } } }
$compInactive = @($compInactive) | Sort-Object 'Jours inactif' -Descending

# ---- Comptes privilégiés (9 groupes) ----
Write-Step "Analyse des comptes privilégiés..."
$privRows = New-Object System.Collections.Generic.List[object]
$privDN = New-Object System.Collections.Generic.HashSet[string]; $privCount = @{}
$privGroups = @(
 @{N='Domain Admins';S="$($domain.DomainSID.Value)-512"},@{N='Enterprise Admins';S="$($domain.DomainSID.Value)-519"},
 @{N='Schema Admins';S="$($domain.DomainSID.Value)-518"},@{N='Group Policy Creator Owners';S="$($domain.DomainSID.Value)-520"},
 @{N='Administrators';S='S-1-5-32-544'},@{N='Account Operators';S='S-1-5-32-548'},@{N='Backup Operators';S='S-1-5-32-551'},
 @{N='Server Operators';S='S-1-5-32-549'},@{N='Print Operators';S='S-1-5-32-550'})
foreach ($pg in $privGroups) {
    try { $g = Get-ADGroup -Identity $pg.S -ErrorAction Stop } catch { continue }
    $m=@(); try { $m=@(Get-ADGroupMember -Identity $g -Recursive -ErrorAction SilentlyContinue | Where-Object { $_.objectClass -eq 'user' }) } catch {}
    $privCount[$pg.N]=$m.Count
    foreach ($u in $m) { [void]$privDN.Add($u.distinguishedName); $o=$userByDN[$u.distinguishedName]
        $privRows.Add([PSCustomObject]@{ 'Groupe'=$pg.N; 'Nom du compte'=$u.SamAccountName; 'Activé'=(if($o){if($o.Enabled){'Activé'}else{'Désactivé'}}else{'N/A'}); "MDP n'expire jamais"=(if($o -and $o.PasswordNeverExpires){'Oui'}else{'Non'}); 'Dernière connexion'=(if($o){$o.LastLogonDate}else{$null}); 'Emplacement (OU)'=(Get-OUFromDN $u.distinguishedName) }) }
}
$privReport = $privRows | Sort-Object 'Groupe','Nom du compte'

# ---- Kerberoastables / MDP à risque ----
$kerb = $users | Where-Object { $_.SamAccountName -ne 'krbtgt' -and @($_.servicePrincipalName).Count -gt 0 } | ForEach-Object {
    [PSCustomObject]@{ 'Nom du compte'=$_.SamAccountName; 'Activé'=(if($_.Enabled){'Activé'}else{'Désactivé'}); "MDP n'expire jamais"=(if($_.PasswordNeverExpires){'Oui'}else{'Non'}); 'SPN'=((@($_.servicePrincipalName)) -join ' ; '); 'Emplacement (OU)'=(Get-OUFromDN $_.DistinguishedName) } } | Sort-Object 'Nom du compte'
$pwdRisk = $users | Where-Object { $_.Enabled -and ($_.PasswordNeverExpires -or $_.PasswordNotRequired -or (-not $_.PasswordLastSet) -or ($_.PasswordLastSet -and (Days-Since $_.PasswordLastSet) -ge $MaxPasswordAgeDays)) } | ForEach-Object {
    [PSCustomObject]@{ 'Nom du compte'=$_.SamAccountName; "MDP n'expire jamais"=(if($_.PasswordNeverExpires){'Oui'}else{'Non'}); 'MDP non requis'=(if($_.PasswordNotRequired){'Oui'}else{'Non'}); 'Âge MDP (jours)'=(Days-Since $_.PasswordLastSet); 'Emplacement (OU)'=(Get-OUFromDN $_.DistinguishedName) } } | Sort-Object 'Âge MDP (jours)' -Descending

# ---- Délégation par flags de compte + SIDHistory ----
$delegAcc = New-Object System.Collections.Generic.List[object]
$sidHist  = New-Object System.Collections.Generic.List[object]
foreach ($o in @($users) + @($computers)) {
    $type = if ($o.TrustedForDelegation) { 'Non contrainte (CRITIQUE)' } elseif ($o.TrustedToAuthForDelegation) { 'Contrainte (transition protocole)' } elseif (@($o.'msDS-AllowedToDelegateTo').Count -gt 0) { 'Contrainte' } else { $null }
    if ($type) { $delegAcc.Add([PSCustomObject]@{ 'Objet'=$o.SamAccountName; "Type d'objet"=(if($o.objectClass -eq 'computer'){'Ordinateur'}else{'Utilisateur'}); 'Type délégation'=$type; 'Cibles'=((@($o.'msDS-AllowedToDelegateTo')) -join ' ; '); 'Emplacement (OU)'=(Get-OUFromDN $o.DistinguishedName) }) }
    if (@($o.SIDHistory).Count -gt 0) { $sidHist.Add([PSCustomObject]@{ 'Objet'=$o.SamAccountName; "Type d'objet"=(if($o.objectClass -eq 'computer'){'Ordinateur'}else{'Utilisateur'}); 'Nb SIDHistory'=@($o.SIDHistory).Count; 'SIDHistory'=((@($o.SIDHistory)) -join ' ; ') }) }
}
$unconstrained = @($delegAcc | Where-Object { $_.'Type délégation' -like 'Non contrainte*' }).Count

# ---- krbtgt / Guest ----
$krbtgt = $users | Where-Object { $_.SamAccountName -eq 'krbtgt' } | Select-Object -First 1
$krbtgtAge = if ($krbtgt) { Days-Since $krbtgt.PasswordLastSet } else { $null }
$guest = $users | Where-Object { $_.SamAccountName -eq 'Guest' } | Select-Object -First 1
$guestOn = if ($guest -and $guest.Enabled) { 'Oui' } else { 'Non' }

# ---- Délégations ACL (DCSync, resetPwd, GenericAll…) ----
$aclRows = New-Object System.Collections.Generic.List[object]
if (-not $SkipAclScan) {
    Write-Step "Analyse des délégations ACL (peut être long)..."
    $G_RESET='00299570-246d-11d0-a768-00aa006e0529'; $G_DC1='1131f6aa-9c07-11d1-f79f-00c04fc2dcd2'; $G_DC2='1131f6ad-9c07-11d1-f79f-00c04fc2dcd2'; $G_MEMBER='bf9679c0-0de6-11d0-a285-00aa003049e2'
    $safe = 'Domain Admins|Admins du domaine|Enterprise Admins|Administrateurs de|Schema Admins|Administrators|Administrateurs|SYSTEM|Système|Enterprise Domain Controllers|Domain Controllers|Contrôleurs de domaine|BUILTIN|NT AUTHORITY|AUTORITE NT|Key Admins|Administrateurs clés|CREATOR OWNER|CREATEUR|Self|Cert Publishers'
    $tg = New-Object System.Collections.Generic.List[string]
    $tg.Add($domain.DistinguishedName); $tg.Add("CN=AdminSDHolder,CN=System,$($domain.DistinguishedName)")
    Get-ADOrganizationalUnit -Filter * -ResultSetSize $null @baseParam | ForEach-Object { $tg.Add($_.DistinguishedName) }
    foreach ($rid in 512,519,518,520) { try { $tg.Add((Get-ADGroup -Identity "$($domain.DomainSID.Value)-$rid" -ErrorAction Stop).DistinguishedName) } catch {} }
    foreach ($dn in ($tg | Select-Object -Unique)) {
        $acl=$null; try { $acl = Get-Acl -Path ("AD:$dn") -ErrorAction Stop } catch { continue }
        foreach ($ace in $acl.Access) {
            if ($ace.AccessControlType -ne 'Allow') { continue }
            $who="$($ace.IdentityReference)"; if ($who -match $safe) { continue }
            $r="$($ace.ActiveDirectoryRights)"; $ot="$($ace.ObjectType)"; $droit=$null; $risque=$null
            if ($ot -eq $G_DC1 -or $ot -eq $G_DC2) { $droit='DCSync'; $risque='CRITIQUE' }
            elseif ($ot -eq $G_RESET -and $r -match 'ExtendedRight') { $droit='Réinitialiser MDP'; $risque='ELEVE' }
            elseif ($r -match 'GenericAll') { $droit='GenericAll'; $risque='ELEVE' }
            elseif ($r -match 'WriteDacl') { $droit='WriteDacl'; $risque='ELEVE' }
            elseif ($r -match 'WriteOwner') { $droit='WriteOwner'; $risque='ELEVE' }
            elseif ($ot -eq $G_MEMBER -and $r -match 'WriteProperty|GenericWrite') { $droit='Écrire membres'; $risque='ELEVE' }
            elseif ($r -match 'GenericWrite') { $droit='GenericWrite'; $risque='À surveiller' }
            else { continue }
            $aclRows.Add([PSCustomObject]@{ 'Objet'=$dn; 'Principal'=$who; 'Droit'=$droit; 'Risque'=$risque })
        }
    }
}
$aclReport = $aclRows | Sort-Object @{E={switch($_.Risque){'CRITIQUE'{0}'ELEVE'{1}default{2}}}},'Objet'
$dcsyncCount = @($aclReport | Where-Object { $_.Droit -eq 'DCSync' }).Count

# ---- Groupes + membres ----
$groupsReport = $groups | ForEach-Object { [PSCustomObject]@{ 'Nom du groupe'=$_.Name; 'Description'=$_.Description; 'Catégorie'="$($_.GroupCategory)"; 'Étendue'="$($_.GroupScope)"; 'Nb membres'=@($_.member).Count; 'Date de création'=$_.whenCreated; 'Emplacement (OU)'=(Get-OUFromDN $_.DistinguishedName) } } | Sort-Object 'Nom du groupe'
$emptyGroupsList = $groupsReport | Where-Object { $_.'Nb membres' -eq 0 }
$grpUserMembers = New-Object System.Collections.Generic.List[object]; $grpCompMembers = New-Object System.Collections.Generic.List[object]
if (-not $SkipGroupMembers) {
    Write-Step "Analyse des membres par groupe..."
    foreach ($g in $groups) {
        foreach ($dn in @($g.member | Where-Object { $_ })) {
            if ($userByDN.ContainsKey($dn)) { $u=$userByDN[$dn]; $grpUserMembers.Add([PSCustomObject]@{ 'Groupe'=$g.Name; 'Nom du compte'=$u.SamAccountName; 'Nom complet'=$u.DisplayName; 'Statut'=(if($u.Enabled){'Activé'}else{'Désactivé'}); 'Dernière connexion'=$u.LastLogonDate; 'Emplacement (OU)'=(Get-OUFromDN $u.DistinguishedName) }) }
            elseif ($compByDN.ContainsKey($dn)) { $c=$compByDN[$dn]; $grpCompMembers.Add([PSCustomObject]@{ 'Groupe'=$g.Name; 'Nom du PC'=$c.Name; 'OS'=$c.OperatingSystem; 'Statut'=(if($c.Enabled){'Activé'}else{'Désactivé'}); 'Obsolète'=(Get-ObsolescenceStatus $c.OperatingSystem); 'Emplacement (OU)'=(Get-OUFromDN $c.DistinguishedName) }) }
        }
    }
}
$grpUserMembers = $grpUserMembers | Sort-Object 'Groupe','Nom du compte'
$grpCompMembers = $grpCompMembers | Sort-Object 'Groupe','Nom du PC'

# ---- Trusts ----
$trusts=@(); try { $trusts=Get-ADTrust -Filter * -Properties Direction,TrustType,IntraForest,SIDFilteringQuarantined,SIDFilteringForestAware,whenCreated,Target } catch {}
$trustReport = $trusts | ForEach-Object { $sid=if($_.SIDFilteringQuarantined -or $_.SIDFilteringForestAware){'Actif'}else{'Désactivé'}; [PSCustomObject]@{ 'Partenaire'=$_.Target; 'Sens'="$($_.Direction)"; 'Type'="$($_.TrustType)"; 'Intra-forêt'=(if($_.IntraForest){'Oui'}else{'Non'}); 'Filtrage SID'=$sid; 'Date de création'=$_.whenCreated } }

# ---- GPO + liens + cpassword ----
$gpoReport=@(); $gpoLinks=New-Object System.Collections.Generic.List[object]; $cpwd=New-Object System.Collections.Generic.List[object]
if (Get-Module -ListAvailable -Name GroupPolicy) {
    Import-Module GroupPolicy -ErrorAction SilentlyContinue
    Write-Step "Analyse des GPO..."
    try {
        foreach ($g in (Get-GPO -All)) {
            $links=@(); $filt=@()
            try { [xml]$rep=Get-GPOReport -Guid $g.Id -ReportType Xml -ErrorAction Stop; if ($rep.GPO.LinksTo){$links=@($rep.GPO.LinksTo)} } catch {}
            try { $filt=@(Get-GPPermission -Guid $g.Id -All -ErrorAction Stop | Where-Object { $_.Permission -eq 'GpoApply' } | ForEach-Object { $_.Trustee.Name }) } catch {}
            $gpoReport += [PSCustomObject]@{ 'Nom de la GPO'=$g.DisplayName; 'Statut'="$($g.GpoStatus)"; 'Date de création'=$g.CreationTime; 'Date de modification'=$g.ModificationTime; 'Nb liens'=@($links).Count; 'Nb entités filtrage'=@($filt).Count; 'Chemin SYSVOL'=("\\{0}\SYSVOL\{0}\Policies\{{{1}}}" -f $g.DomainName,$g.Id) }
            foreach ($l in $links) { $gpoLinks.Add([PSCustomObject]@{ 'GPO'=$g.DisplayName; 'Lien (emplacement)'=$l.SOMPath; 'Activé'=(if("$($l.Enabled)" -eq 'true'){'Oui'}else{'Non'}) }) }
        }
        $gpoReport = $gpoReport | Sort-Object 'Nom de la GPO'
    } catch {}
    try {
        $sysvol = "\\{0}\SYSVOL\{0}\Policies" -f $domain.DNSRoot
        Get-ChildItem -Path $sysvol -Recurse -File -Include *.xml -ErrorAction SilentlyContinue | ForEach-Object {
            if (Select-String -Path $_.FullName -Pattern 'cpassword="([^"]+)"' -ErrorAction SilentlyContinue) { $cpwd.Add([PSCustomObject]@{ 'Fichier'=$_.Name; 'Chemin'=$_.FullName; 'Risque'='CRITIQUE' }) }
        }
    } catch {}
}

# ---- LAPS ----
$schema=(Get-ADRootDSE).schemaNamingContext
function Test-Attr { param($n) try { return ([bool](Get-ADObject -SearchBase $schema -LDAPFilter "(lDAPDisplayName=$n)" -ErrorAction Stop)) } catch { return $false } }
$hasLegacy=Test-Attr 'ms-Mcs-AdmPwdExpirationTime'; $hasNew=Test-Attr 'msLAPS-PasswordExpirationTime'
$lapsReport=@(); $lapsManaged=0
if ($hasLegacy -or $hasNew) {
    $lp=@('Name','Enabled'); if($hasLegacy){$lp+='ms-Mcs-AdmPwdExpirationTime'}; if($hasNew){$lp+='msLAPS-PasswordExpirationTime'}
    try { $lc=Get-ADComputer -Filter 'OperatingSystem -like "*Windows*"' -Properties $lp @baseParam
        $lapsReport = foreach ($c in $lc) { $exp=$null; $t=''
            if ($hasLegacy -and $c.'ms-Mcs-AdmPwdExpirationTime') { $exp=[datetime]::FromFileTime([int64]$c.'ms-Mcs-AdmPwdExpirationTime'); $t='LAPS legacy' }
            if ($hasNew -and $c.'msLAPS-PasswordExpirationTime') { $exp=[datetime]::FromFileTime([int64]$c.'msLAPS-PasswordExpirationTime'); $t='Windows LAPS' }
            [PSCustomObject]@{ 'Nom du PC'=$c.Name; 'LAPS géré'=(if($exp){'Oui'}else{'Non'}); 'Type LAPS'=$t; 'Expiration MDP'=$exp } }
        $lapsManaged = @($lapsReport | Where-Object { $_.'LAPS géré' -eq 'Oui' }).Count
    } catch {}
}
$lapsCov = if (@($computers).Count -gt 0) { [math]::Round(($lapsManaged/@($computers).Count)*100,1) } else { 0 }

# ---- BitLocker ----
$withKey=New-Object System.Collections.Generic.HashSet[string]
try { Get-ADObject -LDAPFilter '(objectClass=msFVE-RecoveryInformation)' -Properties distinguishedName @baseParam | ForEach-Object { [void]$withKey.Add(($_.DistinguishedName -replace '^CN=[^,]+,','')) } } catch {}
$noBitlocker = foreach ($c in $computers) { if (-not $withKey.Contains($c.DistinguishedName)) { [PSCustomObject]@{ 'Nom du PC'=$c.Name; 'OS'=$c.OperatingSystem; 'Statut'=(if($c.Enabled){'Activé'}else{'Désactivé'}); 'Emplacement (OU)'=(Get-OUFromDN $c.DistinguishedName) } } }
$noBitlocker=@($noBitlocker) | Sort-Object 'Nom du PC'
$blCov = if (@($computers).Count -gt 0) { [math]::Round(((@($computers).Count-@($noBitlocker).Count)/@($computers).Count)*100,1) } else { 0 }

# ---- Doublons complets ----
Write-Step "Analyse des doublons..."
$allObjInfos = New-Object System.Collections.Generic.List[object]
foreach ($u in $users)     { $allObjInfos.Add([PSCustomObject]@{ Raw=$u; Name=$u.Name; Class='Utilisateur' }) }
foreach ($c in $computers) { $allObjInfos.Add([PSCustomObject]@{ Raw=$c; Name=$c.Name; Class='Ordinateur' }) }
foreach ($g in $groups)    { $allObjInfos.Add([PSCustomObject]@{ Raw=$g; Name=$g.Name; Class='Groupe' }) }
function Find-Dups2 { param($Selector,[string]$Type)
    $map=@{}; foreach ($i in $allObjInfos) { foreach ($k in (& $Selector $i)) { if([string]::IsNullOrWhiteSpace($k)){continue}; $kk="$k".ToLower().Trim(); if(-not $map.ContainsKey($kk)){$map[$kk]=New-Object System.Collections.Generic.List[object]}; $map[$kk].Add($i.Raw.SamAccountName) } }
    $rows=@(); foreach ($k in $map.Keys) { if (@($map[$k]).Count -gt 1) { $rows += [PSCustomObject]@{ 'Type'=$Type; 'Valeur'=$k; 'Occurrences'=@($map[$k]).Count; 'Objets'=((@($map[$k])) -join ' ; ') } } }
    return $rows
}
$dups=@()
$dups += Find-Dups2 { param($i) @($i.Raw.servicePrincipalName) } 'SPN'
$dups += Find-Dups2 { param($i) ,$i.Raw.UserPrincipalName } 'UPN'
$dups += Find-Dups2 { param($i) $k=@(); if($i.Raw.mail){$k+=$i.Raw.mail}; foreach($p in @($i.Raw.proxyAddresses)){if($p){$k+=($p -replace '^(?i)smtp:','')}}; $k } 'E-mail'
$dups += Find-Dups2 { param($i) ,$i.Raw.DisplayName } 'DisplayName'
$dups += Find-Dups2 { param($i) ,$i.Name } 'Nom (CN)'
$dups += Find-Dups2 { param($i) ,$i.Raw.employeeID } 'employeeID'
# Objets en conflit CNF
try { Get-ADObject -LDAPFilter '(|(cn=*CNF:*)(ou=*CNF:*)(name=*CNF:*))' -Properties objectClass @baseParam | ForEach-Object { $dups += [PSCustomObject]@{ 'Type'='Objet en conflit (CNF)'; 'Valeur'=$_.Name; 'Occurrences'=1; 'Objets'="$($_.objectClass)" } } } catch {}
$dups=@($dups) | Sort-Object 'Type','Occurrences' -Descending

# ---- Échéances + OU/groupes vides ----
$limit=$now.AddDays(30)
$echeances = @()
$echeances += $users | Where-Object { $_.AccountExpirationDate -and $_.AccountExpirationDate -le $limit } | ForEach-Object { [PSCustomObject]@{ 'Type'='Compte'; 'Nom du compte'=$_.SamAccountName; 'Échéance'=$_.AccountExpirationDate } }
$echeances += $users | Where-Object { $_.Enabled -and -not $_.PasswordNeverExpires -and $_.'msDS-UserPasswordExpiryTimeComputed' -and $_.'msDS-UserPasswordExpiryTimeComputed' -ne 0 -and $_.'msDS-UserPasswordExpiryTimeComputed' -ne 9223372036854775807 } | ForEach-Object { $e=[datetime]::FromFileTime([int64]$_.'msDS-UserPasswordExpiryTimeComputed'); if($e -le $limit){ [PSCustomObject]@{ 'Type'='Mot de passe'; 'Nom du compte'=$_.SamAccountName; 'Échéance'=$e } } }
$echeances=@($echeances) | Sort-Object 'Échéance'
Write-Step "Analyse des OU vides..."
$vides = New-Object System.Collections.Generic.List[object]
foreach ($ou in (Get-ADOrganizationalUnit -Filter * -ResultSetSize $null @baseParam)) { $ch=@(Get-ADObject -SearchBase $ou.DistinguishedName -SearchScope OneLevel -Filter * -ErrorAction SilentlyContinue | Where-Object { $_.DistinguishedName -ne $ou.DistinguishedName }); if ($ch.Count -eq 0) { $vides.Add([PSCustomObject]@{ 'Type'='OU vide'; 'Nom'=$ou.Name; 'DN'=$ou.DistinguishedName }) } }
foreach ($g in $emptyGroupsList) { $vides.Add([PSCustomObject]@{ 'Type'='Groupe vide'; 'Nom'=$g.'Nom du groupe'; 'DN'=$g.'Emplacement (OU)' }) }

# ---- FSMO / Santé DC / Réplication ----
Write-Step "Analyse des contrôleurs de domaine..."
$fsmo = @(
 [PSCustomObject]@{ 'Rôle'='Schema Master'; 'Titulaire'=$forest.SchemaMaster },
 [PSCustomObject]@{ 'Rôle'='Domain Naming Master'; 'Titulaire'=$forest.DomainNamingMaster },
 [PSCustomObject]@{ 'Rôle'='PDC Emulator'; 'Titulaire'=$domain.PDCEmulator },
 [PSCustomObject]@{ 'Rôle'='RID Master'; 'Titulaire'=$domain.RIDMaster },
 [PSCustomObject]@{ 'Rôle'='Infrastructure Master'; 'Titulaire'=$domain.InfrastructureMaster })
$dcs = Get-ADDomainController -Filter *
$dcReport = foreach ($dc in $dcs) {
    $ping=$false; try{$ping=Test-Connection -ComputerName $dc.HostName -Count 1 -Quiet -ErrorAction SilentlyContinue}catch{}
    $fails=-1; try{$fails=@(Get-ADReplicationFailure -Target $dc.HostName -ErrorAction Stop).Count}catch{}
    [PSCustomObject]@{ 'Contrôleur'=$dc.HostName; 'Site'=$dc.Site; 'OS'=$dc.OperatingSystem; 'Global Catalog'=(if($dc.IsGlobalCatalog){'Oui'}else{'Non'}); 'Ping'=(if($ping){'OK'}else{'KO'}); 'Échecs réplication'=(if($fails -lt 0){'N/A'}else{$fails}); 'État'=(if(-not $ping){'INJOIGNABLE'}elseif($fails -gt 0){'RÉPLICATION KO'}else{'OK'}) }
}
$repl=@(); foreach ($dc in $dcs) { try { Get-ADReplicationPartnerMetadata -Target $dc.HostName -ErrorAction Stop | ForEach-Object { $repl += [PSCustomObject]@{ 'DC'=$dc.HostName; 'Partenaire'=($_.Partner -replace '^CN=NTDS Settings,CN=',''); 'Dernier succès'=$_.LastReplicationSuccess; 'Nb échecs consécutifs'=$_.ConsecutiveReplicationFailures } } } catch {} }

# ---- AD CS ----
$adcsReport=@()
try {
    $conf=(Get-ADRootDSE).configurationNamingContext; $EKUC=@('1.3.6.1.5.5.7.3.2','1.3.6.1.5.2.3.4','1.3.6.1.4.1.311.20.2.2','2.5.29.37.0')
    $adcsReport = Get-ADObject -SearchBase "CN=Certificate Templates,CN=Public Key Services,CN=Services,$conf" -LDAPFilter '(objectClass=pKICertificateTemplate)' -Properties displayName,'msPKI-Certificate-Name-Flag','msPKI-Enrollment-Flag',pKIExtendedKeyUsage -ErrorAction Stop | ForEach-Object {
        $nf=[int64]($_.'msPKI-Certificate-Name-Flag'); $ef=[int64]($_.'msPKI-Enrollment-Flag'); $ekus=@($_.pKIExtendedKeyUsage)
        $ss=(($nf -band 0x1) -ne 0); $mgr=(($ef -band 0x2) -ne 0); $ca=(($ekus.Count -eq 0) -or [bool]($ekus | Where-Object { $EKUC -contains $_ }))
        [PSCustomObject]@{ 'Modèle'=(if($_.displayName){$_.displayName}else{$_.Name}); 'Sujet fourni'=(if($ss){'Oui'}else{'Non'}); 'Auth client'=(if($ca){'Oui'}else{'Non'}); 'Approbation'=(if($mgr){'Oui'}else{'Non'}); 'Risque'=(if($ss -and $ca -and -not $mgr){'ESC1 possible'}else{'OK'}) } }
} catch {}

# ==========================================================================
# KPI
# ==========================================================================
$L = { param($v,$w,$c) if ($v -ge $c){'Critique'} elseif ($v -ge $w){'À surveiller'} else {'OK'} }
$dash = @(
 @('Comptes','Total utilisateurs',@($users).Count,''),
 @('Comptes','Activés',@($users|Where-Object Enabled).Count,''),
 @('Comptes','Désactivés',@($users|Where-Object{-not $_.Enabled}).Count,''),
 @('Comptes',"Inactifs (> $InactiveDays j)",@($userInactive).Count,(& $L @($userInactive).Count 1 100)),
 @('Comptes',"MDP n'expire jamais (activés)",@($users|Where-Object{$_.PasswordNeverExpires -and $_.Enabled}).Count,(& $L @($users|Where-Object{$_.PasswordNeverExpires -and $_.Enabled}).Count 1 20)),
 @('Comptes','MDP non requis (activés)',@($users|Where-Object{$_.PasswordNotRequired -and $_.Enabled}).Count,(& $L @($users|Where-Object{$_.PasswordNotRequired -and $_.Enabled}).Count 1 1)),
 @('Ordinateurs','Total ordinateurs',@($computers).Count,''),
 @('Ordinateurs','OS obsolètes',@($compReport|Where-Object{$_.Obsolète -eq 'Oui'}).Count,(& $L @($compReport|Where-Object{$_.Obsolète -eq 'Oui'}).Count 1 50)),
 @('Ordinateurs',"Inactifs (> $InactiveDays j)",@($compInactive).Count,(& $L @($compInactive).Count 1 50)),
 @('Ordinateurs','Sans clé BitLocker (AD)',@($noBitlocker).Count,(& $L @($noBitlocker).Count 1 50)),
 @('Ordinateurs','Couverture BitLocker (%)',$blCov,''),
 @('Ordinateurs','Couverture LAPS (%)',$lapsCov,''),
 @('Sécurité','Membres Domain Admins',[int]$privCount['Domain Admins'],(& $L ([int]$privCount['Domain Admins']) 5 10)),
 @('Sécurité','Membres Enterprise Admins',[int]$privCount['Enterprise Admins'],(& $L ([int]$privCount['Enterprise Admins']) 1 3)),
 @('Sécurité','Total comptes privilégiés',$privDN.Count,(& $L $privDN.Count 20 50)),
 @('Sécurité','Kerberoastables',@($kerb).Count,(& $L @($kerb).Count 1 5)),
 @('Sécurité','Délégation non contrainte',$unconstrained,(& $L $unconstrained 1 1)),
 @('Sécurité','Délégations ACL sensibles',@($aclReport).Count,(& $L @($aclReport).Count 1 20)),
 @('Sécurité','DCSync délégué',$dcsyncCount,(& $L $dcsyncCount 1 1)),
 @('Sécurité','Objets avec SIDHistory',@($sidHist).Count,(& $L @($sidHist).Count 1 10)),
 @('Sécurité',"Âge MDP krbtgt (jours)",[int]$krbtgtAge,(& $L ([int]$krbtgtAge) 180 365)),
 @('Sécurité','Compte invité activé',$guestOn,(if($guestOn -eq 'Oui'){'Critique'}else{'OK'})),
 @('Sécurité','Mots de passe GPP (cpassword)',@($cpwd).Count,(& $L @($cpwd).Count 1 1)),
 @('Sécurité','Trusts sans filtrage SID',@($trustReport|Where-Object{$_.'Filtrage SID' -eq 'Désactivé'}).Count,(& $L @($trustReport|Where-Object{$_.'Filtrage SID' -eq 'Désactivé'}).Count 1 1)),
 @('Sécurité','Modèles ADCS ESC1 possible',@($adcsReport|Where-Object{$_.Risque -eq 'ESC1 possible'}).Count,(& $L @($adcsReport|Where-Object{$_.Risque -eq 'ESC1 possible'}).Count 1 1)),
 @('Sécurité','Doublons SPN',@($dups|Where-Object{$_.Type -eq 'SPN'}).Count,(& $L @($dups|Where-Object{$_.Type -eq 'SPN'}).Count 1 1)),
 @('Hygiène','Groupes',@($groups).Count,''),
 @('Hygiène','Groupes vides',@($emptyGroupsList).Count,''),
 @('Hygiène','OU vides',@($vides|Where-Object{$_.Type -eq 'OU vide'}).Count,''),
 @('Hygiène','GPO',@($gpoReport).Count,''),
 @('Hygiène','GPO non liées',@($gpoReport|Where-Object{$_.'Nb liens' -eq 0}).Count,''),
 @('Hygiène','Approbations (trusts)',@($trustReport).Count,''),
 @('Hygiène','Échéances (< 30 j)',@($echeances).Count,(& $L @($echeances).Count 1 100)),
 @('Hygiène','Contrôleurs de domaine',@($dcReport).Count,''),
 @('Hygiène','DC injoignables',@($dcReport|Where-Object{$_.'État' -eq 'INJOIGNABLE'}).Count,(& $L @($dcReport|Where-Object{$_.'État' -eq 'INJOIGNABLE'}).Count 1 1))
)

# ==========================================================================
# EXPORT
# ==========================================================================
Write-Step "Génération du classeur : $OutputPath"
if (Test-Path $OutputPath) { Remove-Item $OutputPath -Force }
$script:excel = $null
function Add-Detail { param($Data,$Sheet,$Table,$DateCols)
    if (@($Data).Count -eq 0) { $Data = ,([PSCustomObject]@{ Information='Aucun élément' }) }
    $script:excel = if ($null -eq $script:excel) { $Data | Export-Excel -Path $OutputPath -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
                    else { $Data | Export-Excel -ExcelPackage $script:excel -WorksheetName $Sheet -AutoSize -AutoFilter -FreezeTopRow -BoldTopRow -TableName $Table -TableStyle 'Medium2' -PassThru }
    if ($DateCols) { $ws=$script:excel.Workbook.Worksheets[$Sheet]; $hr=$ws.Dimension.Start.Row; for($c=1;$c -le $ws.Dimension.End.Column;$c++){ if($DateCols -contains $ws.Cells[$hr,$c].Value){ $l=[OfficeOpenXml.ExcelCellAddress]::GetColumnLetter($c); $ws.Cells[("{0}{1}:{0}{2}" -f $l,($hr+1),$ws.Dimension.End.Row)].Style.Numberformat.Format='dd-mm-yyyy hh:mm' } } }
}
Add-Detail $usersReport    'Utilisateurs'         'Utilisateurs'   @('Dernière connexion','Dernier changement MDP','Date de création')
Add-Detail $compReport     'Ordinateurs'          'Ordinateurs'    @('Date de création','Dernière connexion')
Add-Detail $userInactive   'Comptes inactifs'     'ComptesInactifs' @('Dernière connexion','Date de création')
Add-Detail $compInactive   'Ordinateurs inactifs' 'PCInactifs'     @('Dernière connexion')
Add-Detail $privReport     'Comptes privilégiés'  'Privilegies'    @('Dernière connexion')
Add-Detail $kerb           'Kerberoastables'      'Kerberoast'
Add-Detail $pwdRisk        'MDP à risque'         'MDPRisque'
Add-Detail $delegAcc       'Délégation (comptes)' 'DelegComptes'
Add-Detail $aclReport      'Délégations ACL'      'DelegACL'
Add-Detail $sidHist        'SIDHistory'           'SIDHistory'
Add-Detail $groupsReport   'Groupes'              'Groupes'        @('Date de création')
if (-not $SkipGroupMembers) { Add-Detail $grpUserMembers 'Membres Utilisateurs' 'MembresU' @('Dernière connexion'); Add-Detail $grpCompMembers 'Membres Ordinateurs' 'MembresC' }
Add-Detail $trustReport    'Trusts'               'Trusts'         @('Date de création')
Add-Detail $gpoReport      'GPO'                  'GPO'            @('Date de création','Date de modification')
Add-Detail $gpoLinks       'Liens GPO'            'LiensGPO'
Add-Detail $cpwd           'cpassword (GPP)'      'CPassword'
Add-Detail $lapsReport     'LAPS'                 'LAPS'           @('Expiration MDP')
Add-Detail $noBitlocker    'Sans BitLocker'       'SansBitLocker'
Add-Detail $dups           'Doublons'             'Doublons'
Add-Detail @($vides)       'OU-Groupes vides'     'Vides'
Add-Detail $echeances      'Échéances'            'Echeances'      @('Échéance')
Add-Detail $fsmo           'FSMO'                 'FSMO'
Add-Detail $dcReport       'Santé DC'             'SanteDC'
Add-Detail $repl           'Réplication'          'Replication'    @('Dernier succès')
Add-Detail $adcsReport     'AD CS'                'ADCS'
$excel = $script:excel

# ---- Tableau de bord ----
$wsK = Add-Worksheet -ExcelPackage $excel -WorksheetName 'Tableau de bord (KPI)' -MoveToStart
$wsK.Cells['B2'].Value = 'Rapport Active Directory COMPLET - Tableau de bord'
$wsK.Cells['B2'].Style.Font.Size = 18; $wsK.Cells['B2'].Style.Font.Bold = $true; $wsK.Cells['B2:E2'].Merge = $true
$wsK.Cells['B3'].Value = ("Domaine : {0}  |  Date : {1}" -f $domain.DNSRoot, (Get-Date -Format 'dd-MM-yyyy HH:mm'))
$wsK.Cells['B4'].Value = ("Auteur : {0}" -f $Author)
$colOK=[System.Drawing.Color]::FromArgb(198,239,206); $colW=[System.Drawing.Color]::FromArgb(255,235,156); $colC=[System.Drawing.Color]::FromArgb(255,199,206)
$row=6; $curCat=''
foreach ($k in $dash) {
    if ($k[0] -ne $curCat) { $row+=1; $curCat=$k[0]
        $wsK.Cells["B$row"].Value=$curCat; $wsK.Cells["B$row:D$row"].Merge=$true
        $wsK.Cells["B$row:D$row"].Style.Font.Bold=$true; $wsK.Cells["B$row:D$row"].Style.Font.Size=12
        $wsK.Cells["B$row:D$row"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid
        $wsK.Cells["B$row:D$row"].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::FromArgb(68,114,196))
        $wsK.Cells["B$row:D$row"].Style.Font.Color.SetColor([System.Drawing.Color]::White); $row++ }
    $wsK.Cells["B$row"].Value=$k[1]; $wsK.Cells["C$row"].Value=$k[2]; $wsK.Cells["D$row"].Value=$k[3]
    switch ($k[3]) { 'Critique' { $wsK.Cells["D$row"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["D$row"].Style.Fill.BackgroundColor.SetColor($colC); $wsK.Cells["D$row"].Style.Font.Bold=$true }
                    'À surveiller' { $wsK.Cells["D$row"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["D$row"].Style.Fill.BackgroundColor.SetColor($colW) }
                    'OK' { $wsK.Cells["D$row"].Style.Fill.PatternType=[OfficeOpenXml.Style.ExcelFillStyle]::Solid; $wsK.Cells["D$row"].Style.Fill.BackgroundColor.SetColor($colOK) } }
    $row++
}
$wsK.Column(2).Width=42; $wsK.Column(3).Width=16; $wsK.Column(4).Width=16
$osDist=$compReport | Group-Object OS | Sort-Object Count -Descending | Select-Object -First 8
$fr=7; $wsK.Cells["F6"].Value='Système'; $wsK.Cells["G6"].Value='Nombre'; $wsK.Cells["F6:G6"].Style.Font.Bold=$true
foreach ($o in $osDist) { $wsK.Cells["F$fr"].Value=(if($o.Name){$o.Name}else{'(non renseigné)'}); $wsK.Cells["G$fr"].Value=$o.Count; $fr++ }
try { $ch=$wsK.Drawings.AddChart('osFull',[OfficeOpenXml.Drawing.Chart.eChartType]::BarClustered); $ch.Title.Text='Ordinateurs par OS'; $ch.SetPosition(6,0,8,0); $ch.SetSize(520,300); $null=$ch.Series.Add($wsK.Cells["G7:G$($fr-1)"],$wsK.Cells["F7:F$($fr-1)"]); $ch.Legend.Remove() } catch {}
$nC=@($dash|Where-Object{$_[3] -eq 'Critique'}).Count; $nW=@($dash|Where-Object{$_[3] -eq 'À surveiller'}).Count; $nO=@($dash|Where-Object{$_[3] -eq 'OK'}).Count
$rr=($fr+2); $wsK.Cells["F$rr"].Value='Niveau'; $wsK.Cells["G$rr"].Value='Nombre'; $wsK.Cells["F$rr:G$rr"].Style.Font.Bold=$true
$wsK.Cells["F$($rr+1)"].Value='Critique'; $wsK.Cells["G$($rr+1)"].Value=$nC
$wsK.Cells["F$($rr+2)"].Value='À surveiller'; $wsK.Cells["G$($rr+2)"].Value=$nW
$wsK.Cells["F$($rr+3)"].Value='OK'; $wsK.Cells["G$($rr+3)"].Value=$nO
try { $ch2=$wsK.Drawings.AddChart('riskFull',[OfficeOpenXml.Drawing.Chart.eChartType]::Pie); $ch2.Title.Text='Indicateurs par niveau'; $ch2.SetPosition($rr+5,0,8,0); $ch2.SetSize(420,280); $null=$ch2.Series.Add($wsK.Cells["G$($rr+1):G$($rr+3)"],$wsK.Cells["F$($rr+1):F$($rr+3)"]); $ch2.DataLabel.ShowValue=$true } catch {}

Close-ExcelPackage $excel
Write-Host ""
Write-Host ("[OK] Rapport COMPLET généré : {0}" -f $OutputPath) -ForegroundColor Green
Write-Host ("     Indicateurs -> Critique: {0} | À surveiller: {1} | OK: {2}" -f $nC,$nW,$nO) -ForegroundColor Green
