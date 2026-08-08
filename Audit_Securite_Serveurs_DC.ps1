#Requires -Version 5.1
<#
.SYNOPSIS
    Audit de sécurité des partages réseau SMB/NTFS sur les serveurs Windows de l'AD.
    VERSION DC : collecte uniquement, exporte un CSV. Pas de Python requis.

.DESCRIPTION
    Ce script s'exécute sur le Contrôleur de Domaine (ou toute machine avec RSAT).
    Il audite les droits SMB et NTFS, génère un fichier CSV et un fichier JSON
    de métadonnées pour la mise en forme ultérieure sur un poste local.

.PARAMETER OutputDir
    Dossier de sortie. Par défaut : Bureau de l'utilisateur.

.PARAMETER BatchSize
    Nombre de serveurs par vague WinRM. Par défaut : 20.

.PARAMETER ThrottleLimit
    Connexions WinRM simultanées max par vague. Par défaut : 10.

.PARAMETER PauseSec
    Pause réseau en secondes entre les vagues. Par défaut : 5.

.PARAMETER SearchBase
    OU Active Directory à cibler (ex: "OU=Serveurs,DC=example,DC=local").
    Par défaut : tout le domaine.

.PARAMETER IncludeSystemShares
    Inclut les partages administratifs (ADMIN$, C$...). Par défaut : exclus.

.EXAMPLE
    PowerShell.exe -ExecutionPolicy Bypass -File ".\Audit_Securite_Serveurs_DC.ps1"
    .\Audit_Securite_Serveurs_DC.ps1 -OutputDir "C:\Audits" -BatchSize 10 -PauseSec 10
    .\Audit_Securite_Serveurs_DC.ps1 -SearchBase "OU=Serveurs,DC=example,DC=local"

.NOTES
    Version : 3.1 - Edition DC (sans Python)
    Droits requis : Lecture AD + accès WinRM sur les serveurs cibles.
                    Exécution en Administrateur de domaine recommandée.
    Sortie      : Audit_Securite_AAAAMMJJ_HHMM.csv
                  Audit_Securite_AAAAMMJJ_HHMM_meta.json
    Étape suivante : Copier ces deux fichiers sur votre PC local et exécuter
                     Formater_Rapport_Excel.ps1 pour générer le .xlsx
#>

[CmdletBinding()]
param(
    [string] $OutputDir        = "$env:USERPROFILE\Desktop",
    [int]    $BatchSize        = 20,
    [int]    $ThrottleLimit    = 10,
    [int]    $PauseSec         = 5,
    [string] $SearchBase       = "",
    [switch] $IncludeSystemShares
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
#region INITIALISATION & JOURNALISATION
# ==============================================================================

$ScriptVersion = "3.1-DC"
$Date          = Get-Date -Format 'yyyyMMdd_HHmm'
$LogFile       = "$OutputDir\Audit_Securite_$Date.log"
$CsvFile       = "$OutputDir\Audit_Securite_$Date.csv"
$MetaFile      = "$OutputDir\Audit_Securite_$Date`_meta.json"
$ResumeFile    = "$env:TEMP\AuditResume_$($env:USERNAME).json"
$StartTime     = Get-Date

if (!(Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

function Write-Log {
    param(
        [string] $Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')] [string] $Level = 'INFO'
    )
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogLine   = "[$Timestamp] [$Level] $Message"
    Add-Content -Path $LogFile -Value $LogLine -Encoding UTF8 -ErrorAction SilentlyContinue
    $FgColor = switch ($Level) {
        'INFO'    { 'Cyan'     }
        'WARN'    { 'Yellow'   }
        'ERROR'   { 'Red'      }
        'SUCCESS' { 'Green'    }
        'DEBUG'   { 'DarkGray' }
    }
    Write-Host $LogLine -ForegroundColor $FgColor
}

function Write-Section {
    param([string]$Title, [string]$Step)
    Write-Host "`n  $(('-' * 62))" -ForegroundColor DarkBlue
    Write-Host "  $Step  $Title" -ForegroundColor White
    Write-Host "  $(('-' * 62))" -ForegroundColor DarkBlue
    Write-Log "$Step $Title" -Level INFO
}

Write-Host @"

  ╔══════════════════════════════════════════════════════════════╗
  ║      AUDIT SECURITE - PARTAGES RESEAU  v$ScriptVersion           ║
  ║      SMB / NTFS / Active Directory  [EDITION DC]            ║
  ╚══════════════════════════════════════════════════════════════╝
  Utilisateur : $env:USERDOMAIN\$env:USERNAME
  Machine     : $env:COMPUTERNAME
  Date        : $(Get-Date -Format 'dd/MM/yyyy a HH:mm:ss')
  Journal     : $LogFile
  CSV sortie  : $CsvFile
  Metadata    : $MetaFile
"@ -ForegroundColor White

Write-Log "Script v$ScriptVersion demarre. OutputDir=$OutputDir BatchSize=$BatchSize ThrottleLimit=$ThrottleLimit PauseSec=$PauseSec SearchBase='$SearchBase'" -Level INFO

#endregion

# ==============================================================================
#region ETAPE 0 - VERIFICATION DES PREREQUIS
# ==============================================================================

Write-Section "Verification des prerequis" "[0/5]"

# 0.1 Version PowerShell
if ($PSVersionTable.PSVersion.Major -lt 5) {
    Write-Log "PowerShell 5.1 minimum requis (actuel : $($PSVersionTable.PSVersion))" -Level ERROR
    exit 1
}
Write-Log "PowerShell $($PSVersionTable.PSVersion) OK" -Level SUCCESS

# 0.2 Droits administrateur
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Log "Non execute en Administrateur. Certaines lectures NTFS peuvent etre limitees." -Level WARN
} else {
    Write-Log "Execution en mode Administrateur OK" -Level SUCCESS
}

# 0.3 Module Active Directory (RSAT)
Write-Log "Verification du module ActiveDirectory..." -Level INFO
if (!(Get-Module -ListAvailable -Name ActiveDirectory -ErrorAction SilentlyContinue)) {
    Write-Log "Module ActiveDirectory absent. Tentative d installation automatique..." -Level WARN
    try {
        $OS = Get-CimInstance Win32_OperatingSystem -ErrorAction Stop
        if ($OS.ProductType -eq 1) {
            Add-WindowsCapability -Online -Name "Rsat.ActiveDirectory.DS-LDS.Tools~~~~0.0.1.0" -ErrorAction Stop | Out-Null
        } else {
            Install-WindowsFeature -Name RSAT-AD-PowerShell -IncludeManagementTools -ErrorAction Stop | Out-Null
        }
        Import-Module ActiveDirectory -ErrorAction Stop
        Write-Log "Module ActiveDirectory installe et charge avec succes." -Level SUCCESS
    } catch {
        Write-Log "Echec installation ActiveDirectory : $_" -Level ERROR
        Write-Log "Solution : Parametres > Applications > Fonctionnalites facultatives > RSAT: AD DS" -Level WARN
        exit 1
    }
} else {
    Import-Module ActiveDirectory -ErrorAction SilentlyContinue
    Write-Log "Module ActiveDirectory disponible OK" -Level SUCCESS
}

# 0.4 Test de connectivite AD
try {
    $DomainInfo = Get-ADDomain -ErrorAction Stop
    Write-Log "Domaine AD : $($DomainInfo.DNSRoot) - Connecte OK" -Level SUCCESS
} catch {
    Write-Log "Impossible de joindre le domaine Active Directory : $_" -Level ERROR
    exit 1
}

#endregion

# ==============================================================================
#region ETAPE 1 - CIBLAGE DES SERVEURS DANS L AD
# ==============================================================================

Write-Section "Recherche des serveurs Windows dans l Active Directory" "[1/5]"

try {
    $ADParams = @{
        Filter     = 'OperatingSystem -like "*Windows Server*"'
        Properties = 'Name','OperatingSystem','OperatingSystemVersion','IPv4Address','Enabled','LastLogonDate'
    }
    if ($SearchBase -ne "") { $ADParams['SearchBase'] = $SearchBase }

    $AllADComputers = Get-ADComputer @ADParams | Where-Object { $_.Enabled -eq $true }
    $AllServers     = @($AllADComputers | Select-Object -ExpandProperty Name)

    Write-Log "$($AllServers.Count) serveurs Windows actifs trouves dans l AD." -Level SUCCESS
    $AllADComputers | Group-Object OperatingSystem | Sort-Object Count -Descending |
        ForEach-Object { Write-Log "  OS: $($_.Name) - $($_.Count) serveur(s)" -Level DEBUG }
} catch {
    Write-Log "Erreur lors de la requete Active Directory : $_" -Level ERROR
    exit 1
}

if ($AllServers.Count -eq 0) {
    Write-Log "Aucun serveur Windows trouve dans l AD. Fin du script." -Level WARN
    exit 0
}

#endregion

# ==============================================================================
#region ETAPE 2 - FILTRAGE PING + WINRM
# ==============================================================================

Write-Section "Test de connectivite (Ping + WinRM port 5985)" "[2/5]"

$OnlineServers  = [System.Collections.Generic.List[string]]::new()
$OfflineServers = [System.Collections.Generic.List[string]]::new()
$WinRMKO        = [System.Collections.Generic.List[string]]::new()
$PingTotal      = $AllServers.Count
$PingDone       = 0

foreach ($Srv in $AllServers) {
    $PingDone++
    Write-Progress -Activity "Test de connectivite" `
                   -Status "Ping $Srv ($PingDone/$PingTotal)" `
                   -PercentComplete ([int](($PingDone / $PingTotal) * 100))

    $PingOK = $false
    try { $PingOK = Test-Connection -ComputerName $Srv -Count 2 -Quiet -ErrorAction SilentlyContinue }
    catch { $PingOK = $false }

    if ($PingOK) {
        $WinRMOK = $false
        try {
            $TCP   = New-Object System.Net.Sockets.TcpClient
            $Async = $TCP.BeginConnect($Srv, 5985, $null, $null)
            $Wait  = $Async.AsyncWaitHandle.WaitOne(2000, $false)
            if ($Wait -and $TCP.Connected) { $WinRMOK = $true }
            $TCP.Close()
        } catch { $WinRMOK = $false }

        if ($WinRMOK) {
            $OnlineServers.Add($Srv)
        } else {
            Write-Log "  $Srv : Ping OK mais WinRM inaccessible (port 5985)" -Level DEBUG
            $WinRMKO.Add($Srv)
            $OfflineServers.Add($Srv)
        }
    } else {
        $OfflineServers.Add($Srv)
        Write-Log "  $Srv : Hors ligne (ping KO)" -Level DEBUG
    }
}
Write-Progress -Activity "Test de connectivite" -Completed

Write-Log "En ligne + WinRM OK  : $($OnlineServers.Count) serveurs" -Level SUCCESS
Write-Log "Hors ligne / WinRM KO: $($OfflineServers.Count) serveurs" -Level WARN
if ($WinRMKO.Count -gt 0) {
    Write-Log "Serveurs WinRM KO ($($WinRMKO.Count)): $($WinRMKO -join ', ')" -Level WARN
    Write-Log "Conseil : Verifiez que WinRM est actif (winrm quickconfig) et que le pare-feu autorise le port 5985." -Level INFO
}

if ($OnlineServers.Count -eq 0) {
    Write-Log "Aucun serveur joignable avec WinRM. Fin du script." -Level ERROR
    exit 0
}

#endregion

# ==============================================================================
#region ETAPE 3 - SCRIPTBLOCK D ANALYSE DISTANTE
# ==============================================================================

$ScriptBlock = {
    param([bool]$IncludeSys)

    $FoundOnServer = [System.Collections.Generic.List[PSCustomObject]]::new()
    $Errors        = [System.Collections.Generic.List[string]]::new()

    try {
        $AllShares = Get-SmbShare -ErrorAction Stop
    } catch {
        return [PSCustomObject]@{
            Results = [System.Collections.Generic.List[PSCustomObject]]::new()
            Errors  = @("Get-SmbShare a echoue sur $env:COMPUTERNAME : $_")
        }
    }

    $SysPat = '^(IPC\$|ADMIN\$|[A-Z]\$|NETLOGON|SYSVOL|print\$|prnproc\$|FAX\$|CertEnroll)$'
    $Shares = if ($IncludeSys) { $AllShares } else {
        $AllShares | Where-Object { $_.Name -notmatch $SysPat -and $_.Path -notlike 'C:\Windows*' }
    }

    foreach ($Share in $Shares) {
        $Path      = $Share.Path
        $ShareName = $Share.Name

        $ShareACL          = $null
        $ShareEveryone     = $null
        $ShareEveryonePerm = 'N/A'
        try {
            $ShareACL      = Get-SmbShareAccess -Name $ShareName -ErrorAction Stop
            $ShareEveryone = $ShareACL | Where-Object { $_.AccountName -match 'Everyone|Tout le monde' }
            if ($ShareEveryone) {
                $ShareEveryonePerm = ($ShareEveryone |
                    ForEach-Object { "$($_.AccessControlType) ($($_.AccessRight))" }) -join ' / '
            }
        } catch {
            $Errors.Add("$ShareName : erreur lecture ACL SMB - $_")
        }

        $NTFSEveryone     = $null
        $NTFSEveryonePerm = 'N/A'
        $FullControlStr   = 'Aucun'
        $NTFSErreur       = $false
        $NTFSOwner        = 'N/A'

        if ($Path -and (Test-Path $Path -ErrorAction SilentlyContinue)) {
            try {
                $PathACL   = Get-Acl -Path $Path -ErrorAction Stop
                $NTFSOwner = $PathACL.Owner

                $NTFSEveryone = $PathACL.Access |
                    Where-Object { $_.IdentityReference -match 'Everyone|Tout le monde' }

                if ($NTFSEveryone) {
                    $NTFSEveryonePerm = ($NTFSEveryone |
                        ForEach-Object { "$($_.FileSystemRights) [$($_.AccessControlType)]" }) -join ' | '
                }

                $ExclPat = 'Everyone|Tout le monde|NT AUTHORITY\\SYSTEM|Administrateurs|' +
                           'Administrators|BUILTIN|Creator Owner|CREATEUR|NT SERVICE'
                $FCList = $PathACL.Access | Where-Object {
                    ($_.FileSystemRights -band [System.Security.AccessControl.FileSystemRights]::FullControl) -eq
                        [System.Security.AccessControl.FileSystemRights]::FullControl -and
                    $_.IdentityReference -notmatch $ExclPat -and
                    $_.AccessControlType -eq 'Allow'
                } | Select-Object -ExpandProperty IdentityReference | Select-Object -Unique

                if ($FCList) { $FullControlStr = ($FCList -join ', ') }

            } catch {
                $NTFSErreur       = $true
                $NTFSEveryonePerm = "ERREUR: $($_.Exception.Message -replace '\r?\n',' ')"
                $Errors.Add("$ShareName ($Path) : erreur NTFS - $_")
            }
        } else {
            $NTFSEveryonePerm = "Chemin inaccessible ou inexistant"
        }

        $ShareExpose = $null -ne $ShareEveryone
        $NTFSExpose  = $null -ne $NTFSEveryone

        if (-not $ShareExpose -and -not $NTFSExpose) { continue }

        $NiveauRisque = if ($ShareExpose -and $NTFSExpose -and $NTFSEveryonePerm -match 'FullControl') {
            'CRITIQUE'
        } elseif ($ShareExpose -and $NTFSExpose) {
            'ELEVE'
        } elseif ($ShareExpose) {
            'MOYEN'
        } else {
            'FAIBLE'
        }

        $FolderSizeMB = 'N/A'
        try {
            $ItemSample = Get-ChildItem $Path -Recurse -ErrorAction SilentlyContinue -Force |
                          Select-Object -First 1001
            if ($ItemSample.Count -le 1000) {
                $Bytes = ($ItemSample | Where-Object { -not $_.PSIsContainer } |
                    Measure-Object Length -Sum -ErrorAction SilentlyContinue).Sum
                if ($null -ne $Bytes) { $FolderSizeMB = [Math]::Round($Bytes / 1MB, 1) }
            } else {
                $FolderSizeMB = ">1000 fichiers"
            }
        } catch { $FolderSizeMB = 'N/A' }

        $FoundOnServer.Add([PSCustomObject]@{
            Serveur               = $env:COMPUTERNAME
            NomPartage            = $ShareName
            Description           = if ($Share.Description) { $Share.Description } else { '' }
            CheminLocal           = $Path
            ProprietaireNTFS      = $NTFSOwner
            NiveauRisque          = $NiveauRisque
            DroitPartageEveryone  = if ($ShareExpose) { $ShareEveryonePerm } else { 'Non expose' }
            DroitNTFSEveryone     = if ($NTFSExpose)  { $NTFSEveryonePerm  } else { 'Non expose' }
            AutresControleTotal   = $FullControlStr
            TailleDossierMB       = $FolderSizeMB
            LimitUtilisateurs     = if ($Share.ConcurrentUserLimit -eq 0) { 'Illimite' } else { $Share.ConcurrentUserLimit }
            ErreurNTFS            = if ($NTFSErreur) { 'OUI' } else { 'NON' }
            DateAudit             = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
        })
    }

    return [PSCustomObject]@{ Results = $FoundOnServer; Errors = $Errors }
}

#endregion

# ==============================================================================
#region ETAPE 4 - AUDIT PAR VAGUES AVEC REPRISE AUTOMATIQUE
# ==============================================================================

Write-Section "Audit des partages par vagues" "[3/5]"

$AlreadyAudited = [System.Collections.Generic.HashSet[string]]::new()
$Results        = [System.Collections.Generic.List[PSCustomObject]]::new()
$AllAuditErrors = [System.Collections.Generic.List[string]]::new()

if (Test-Path $ResumeFile) {
    try {
        $RD = Get-Content $ResumeFile -Raw | ConvertFrom-Json
        if ($RD.Date -and (New-TimeSpan -Start ([datetime]$RD.Date) -End (Get-Date)).TotalHours -lt 24) {
            Write-Log "Fichier de reprise trouve ($ResumeFile). Chargement..." -Level WARN
            $RD.AuditedServers | ForEach-Object { $AlreadyAudited.Add($_) | Out-Null }
            $RD.Results | ForEach-Object { $Results.Add([PSCustomObject]$_) }
            Write-Log "$($AlreadyAudited.Count) serveurs deja audites, $($Results.Count) resultats charges." -Level INFO
        } else {
            Remove-Item $ResumeFile -Force -ErrorAction SilentlyContinue
        }
    } catch {
        Write-Log "Fichier de reprise invalide, demarrage a zero." -Level WARN
        Remove-Item $ResumeFile -Force -ErrorAction SilentlyContinue
    }
}

$ServersToAudit  = @($OnlineServers | Where-Object { -not $AlreadyAudited.Contains($_) })
$TotalBatches    = [Math]::Ceiling($ServersToAudit.Count / $BatchSize)
$BatchNum        = 0
$AuditedServers  = [System.Collections.Generic.List[string]]($AlreadyAudited)
$ErrorServers    = [System.Collections.Generic.List[string]]::new()

Write-Log "Serveurs a auditer : $($ServersToAudit.Count) (dont $($AlreadyAudited.Count) deja traites en reprise)" -Level INFO

for ($i = 0; $i -lt $ServersToAudit.Count; $i += $BatchSize) {
    $BatchNum++
    $End          = [Math]::Min($i + $BatchSize - 1, $ServersToAudit.Count - 1)
    $CurrentBatch = @($ServersToAudit[$i..$End])

    Write-Log "Vague $BatchNum/$TotalBatches : $($CurrentBatch.Count) hotes ($($CurrentBatch -join ', '))" -Level INFO
    Write-Progress -Activity "Audit des partages" `
                   -Status "Vague $BatchNum / $TotalBatches" `
                   -PercentComplete ([int](($BatchNum / [Math]::Max($TotalBatches,1)) * 100))

    $RetryCount = 0
    $MaxRetries = 2
    $BatchOK    = $false

    while (-not $BatchOK -and $RetryCount -le $MaxRetries) {
        try {
            $BatchRaw = Invoke-Command `
                -ComputerName $CurrentBatch `
                -ScriptBlock $ScriptBlock `
                -ArgumentList ([bool]$IncludeSystemShares) `
                -ThrottleLimit $ThrottleLimit `
                -ErrorAction SilentlyContinue

            foreach ($BR in $BatchRaw) {
                if ($BR -and $BR.PSObject.Properties['Results']) {
                    foreach ($r in $BR.Results) { $Results.Add($r) }
                    foreach ($e in $BR.Errors)  { $AllAuditErrors.Add("[$($BR.PSComputerName)] $e") }
                } elseif ($BR) {
                    $Results.Add($BR)
                }
            }

            foreach ($Srv in $CurrentBatch) { $AuditedServers.Add($Srv) }
            $BatchOK = $true

        } catch {
            $RetryCount++
            if ($RetryCount -le $MaxRetries) {
                Write-Log "Erreur vague $BatchNum (essai $RetryCount/$MaxRetries) : $_. Nouvelle tentative dans 10s..." -Level WARN
                Start-Sleep -Seconds 10
            } else {
                Write-Log "Abandon vague $BatchNum apres $MaxRetries essais : $_" -Level ERROR
                foreach ($Srv in $CurrentBatch) { $ErrorServers.Add($Srv) }
            }
        }
    }

    try {
        [PSCustomObject]@{
            Date           = (Get-Date -Format 'o')
            AuditedServers = @($AuditedServers)
            Results        = @($Results)
        } | ConvertTo-Json -Depth 10 -Compress |
            Set-Content -Path $ResumeFile -Encoding UTF8 -Force
    } catch {
        Write-Log "Impossible d ecrire le fichier de reprise : $_" -Level DEBUG
    }

    if ($i + $BatchSize -lt $ServersToAudit.Count) {
        Write-Log "Pause reseau de $PauseSec secondes..." -Level DEBUG
        Start-Sleep -Seconds $PauseSec
    }
}

Write-Progress -Activity "Audit des partages" -Completed

if ($AllAuditErrors.Count -gt 0) {
    Write-Log "Erreurs WinRM/ACL ($($AllAuditErrors.Count) au total) :" -Level WARN
    $AllAuditErrors | Select-Object -First 30 | ForEach-Object { Write-Log "  $_" -Level DEBUG }
}

if ($ErrorServers.Count -eq 0) {
    Remove-Item $ResumeFile -Force -ErrorAction SilentlyContinue
}

Write-Log "Audit termine. $($Results.Count) partage(s) expose(s) detecte(s) sur $($AuditedServers.Count) serveurs audites." -Level SUCCESS

#endregion

# ==============================================================================
#region ETAPE 5 - EXPORT CSV + FICHIER METADONNEES
# ==============================================================================

Write-Section "Generation du rapport CSV et des metadonnees" "[4/5]"

$ColOrder = @('Serveur','NomPartage','Description','CheminLocal','ProprietaireNTFS',
              'NiveauRisque','DroitPartageEveryone','DroitNTFSEveryone',
              'AutresControleTotal','TailleDossierMB','LimitUtilisateurs',
              'ErreurNTFS','DateAudit')

$SortedResults = if ($Results.Count -gt 0) {
    $Results | Select-Object $ColOrder | Sort-Object @{E='NiveauRisque';A=$true}, Serveur
} else {
    [PSCustomObject]@{
        Serveur = 'AUCUN PARTAGE EXPOSE'; NomPartage = 'N/A'; Description = 'Audit propre'
        CheminLocal = 'N/A'; ProprietaireNTFS = 'N/A'; NiveauRisque = 'N/A'
        DroitPartageEveryone = 'N/A'; DroitNTFSEveryone = 'N/A'; AutresControleTotal = 'N/A'
        TailleDossierMB = 'N/A'; LimitUtilisateurs = 'N/A'; ErreurNTFS = 'NON'
        DateAudit = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
    }
}

$SortedResults | Export-Csv -Path $CsvFile -NoTypeInformation -Encoding UTF8 -Delimiter ";"
Write-Log "CSV genere : $CsvFile ($($Results.Count) ligne(s))" -Level SUCCESS

# -- Fichier de metadonnees JSON (pour le script de mise en forme) --
$ElapsedMin    = [Math]::Round((New-TimeSpan -Start $StartTime -End (Get-Date)).TotalMinutes, 1)

$MetaData = [PSCustomObject]@{
    ScriptVersion  = $ScriptVersion
    DateAudit      = (Get-Date -Format 'dd/MM/yyyy HH:mm:ss')
    Domaine        = $DomainInfo.DNSRoot
    DureeMinutes   = $ElapsedMin
    AllSrvCount    = $AllServers.Count
    OnlineCount    = $OnlineServers.Count
    OfflineCount   = $OfflineServers.Count
    ErrorSrvCount  = $ErrorServers.Count
    PartagesExposes = $Results.Count
    OfflineServers = @($OfflineServers)
    WinRMKO        = @($WinRMKO)
    ErrorServers   = @($ErrorServers)
    AuditErrors    = @($AllAuditErrors | Select-Object -First 200)
    CsvFile        = $CsvFile
}

$MetaData | ConvertTo-Json -Depth 5 |
    Set-Content -Path $MetaFile -Encoding UTF8 -Force
Write-Log "Metadonnees generees : $MetaFile" -Level SUCCESS

#endregion

# ==============================================================================
#region RESUME FINAL
# ==============================================================================

Write-Section "Resume de l execution" "[5/5]"

$Duration   = New-TimeSpan -Start $StartTime -End (Get-Date)
$RiskGroups = $Results | Group-Object NiveauRisque | Sort-Object Name
$RiskLine   = if ($RiskGroups) { ($RiskGroups | ForEach-Object { "$($_.Name): $($_.Count)" }) -join '  |  ' } else { 'N/A' }

Write-Host @"

  +--------------------------------------------------------------+
  |                    RESULTAT DE L AUDIT                       |
  +--------------------------------------------------------------+
  | Duree totale        : $("{0:D2}m {1:D2}s" -f [int]$Duration.TotalMinutes, $Duration.Seconds)
  | Serveurs dans l AD  : $($AllServers.Count)
  | Joignables (WinRM)  : $($OnlineServers.Count)
  | Hors ligne / KO     : $($OfflineServers.Count)
  | Erreurs d audit     : $($ErrorServers.Count)
  | Partages exposes    : $($Results.Count)
  | Niveaux de risque   : $RiskLine
  +--------------------------------------------------------------+
  | CSV genere          : $CsvFile
  | Metadonnees         : $MetaFile
  | Journal             : $LogFile
  +--------------------------------------------------------------+

  ETAPE SUIVANTE :
  Copiez ces 2 fichiers sur votre PC local :
    - $CsvFile
    - $MetaFile
  Puis executez : Formater_Rapport_Excel.ps1

"@ -ForegroundColor White

Write-Log "Termine en $("{0:D2}m {1:D2}s" -f [int]$Duration.TotalMinutes, $Duration.Seconds). Partages exposes: $($Results.Count). CSV: $CsvFile" -Level SUCCESS

if ($ErrorServers.Count -gt 0) {
    Write-Log "Serveurs non audites (erreurs WinRM): $($ErrorServers -join ', ')" -Level WARN
    Write-Log "Action requise: winrm quickconfig sur ces serveurs + verifier pare-feu port 5985." -Level INFO
}

Write-Log "Fin du script DC." -Level INFO

#endregion
