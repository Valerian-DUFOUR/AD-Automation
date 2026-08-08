#Requires -Version 5.1
<#
.SYNOPSIS
    Mise en forme du rapport d'audit de sécurité SMB/NTFS.
    VERSION PC LOCAL : lit le CSV + JSON de métadonnées, génère le fichier Excel .xlsx.

.DESCRIPTION
    Ce script s'exécute sur VOTRE PC LOCAL (pas sur le DC).
    Il nécessite Python 3.7+ et le module openpyxl.
    Il installe automatiquement Python (via winget) et openpyxl si absents.

    Fichiers requis en entrée (produits par Audit_Securite_Serveurs_DC.ps1) :
      - Audit_Securite_AAAAMMJJ_HHMM.csv
      - Audit_Securite_AAAAMMJJ_HHMM_meta.json

.PARAMETER CsvFile
    Chemin complet vers le fichier CSV produit par le script DC.
    Si non fourni, le script cherche le CSV le plus récent sur le Bureau.

.PARAMETER MetaFile
    Chemin complet vers le fichier JSON de métadonnées.
    Si non fourni, déduit automatiquement depuis le nom du CSV.

.PARAMETER OutputDir
    Dossier de sortie pour le .xlsx. Par défaut : même dossier que le CSV.

.EXAMPLE
    .\Formater_Rapport_Excel.ps1
    .\Formater_Rapport_Excel.ps1 -CsvFile "C:\Audits\Audit_Securite_20250115_1430.csv"
    .\Formater_Rapport_Excel.ps1 -CsvFile "C:\Audits\Audit_Securite_20250115_1430.csv" -OutputDir "C:\Rapports"

.NOTES
    Version : 3.1 - Edition PC Local
    Prérequis : Python 3.7+ (installé automatiquement si absent via winget)
                openpyxl (installé automatiquement si absent via pip)
#>

[CmdletBinding()]
param(
    [string] $CsvFile   = "",
    [string] $MetaFile  = "",
    [string] $OutputDir = ""
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ==============================================================================
#region INITIALISATION
# ==============================================================================

$ScriptVersion = "3.1-PC"
$StartTime     = Get-Date

function Write-Log {
    param(
        [string] $Message,
        [ValidateSet('INFO','WARN','ERROR','SUCCESS','DEBUG')] [string] $Level = 'INFO'
    )
    $Timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
    $LogLine   = "[$Timestamp] [$Level] $Message"
    $FgColor = switch ($Level) {
        'INFO'    { 'Cyan'     }
        'WARN'    { 'Yellow'   }
        'ERROR'   { 'Red'      }
        'SUCCESS' { 'Green'    }
        'DEBUG'   { 'DarkGray' }
    }
    Write-Host $LogLine -ForegroundColor $FgColor
}

Write-Host @"

  ╔══════════════════════════════════════════════════════════════╗
  ║      MISE EN FORME RAPPORT AUDIT  v$ScriptVersion              ║
  ║      CSV → Excel (.xlsx) avec mise en forme complète         ║
  ╚══════════════════════════════════════════════════════════════╝
  Utilisateur : $env:USERDOMAIN\$env:USERNAME
  Machine     : $env:COMPUTERNAME
  Date        : $(Get-Date -Format 'dd/MM/yyyy a HH:mm:ss')
"@ -ForegroundColor White

#endregion

# ==============================================================================
#region ETAPE 1 - LOCALISATION DES FICHIERS D ENTREE
# ==============================================================================

Write-Log "Recherche des fichiers d entree..." -Level INFO

# -- Recherche automatique du CSV le plus recent sur le Bureau si non fourni --
if ($CsvFile -eq "") {
    $Desktop = "$env:USERPROFILE\Desktop"
    $Found   = Get-ChildItem -Path $Desktop -Filter "Audit_Securite_*.csv" -ErrorAction SilentlyContinue |
               Sort-Object LastWriteTime -Descending | Select-Object -First 1
    if ($Found) {
        $CsvFile = $Found.FullName
        Write-Log "CSV trouve automatiquement : $CsvFile" -Level INFO
    } else {
        Write-Log "Aucun fichier CSV Audit_Securite_*.csv trouve sur le Bureau." -Level ERROR
        Write-Log "Utilisez le parametre -CsvFile pour specifier le chemin." -Level WARN
        Write-Host "`nUsage : .\Formater_Rapport_Excel.ps1 -CsvFile `"C:\chemin\vers\audit.csv`"`n" -ForegroundColor Yellow
        exit 1
    }
}

if (!(Test-Path $CsvFile)) {
    Write-Log "Fichier CSV introuvable : $CsvFile" -Level ERROR
    exit 1
}

# -- Recherche du fichier de metadonnees (meme nom, suffixe _meta.json) --
if ($MetaFile -eq "") {
    $MetaFile = $CsvFile -replace '\.csv$', '_meta.json'
}

$MetaLoaded = $false
$Meta       = $null

if (Test-Path $MetaFile) {
    try {
        $Meta       = Get-Content $MetaFile -Raw | ConvertFrom-Json
        $MetaLoaded = $true
        Write-Log "Metadonnees chargees : $MetaFile" -Level SUCCESS
    } catch {
        Write-Log "Impossible de lire le fichier de metadonnees ($MetaFile) : $_. Valeurs par defaut utilisees." -Level WARN
    }
} else {
    Write-Log "Fichier de metadonnees introuvable ($MetaFile). Valeurs par defaut utilisees." -Level WARN
}

# -- Dossier de sortie --
if ($OutputDir -eq "") {
    $OutputDir = Split-Path -Parent $CsvFile
}

if (!(Test-Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null
}

$BaseName  = [System.IO.Path]::GetFileNameWithoutExtension($CsvFile)
$XlsxFile  = "$OutputDir\$BaseName.xlsx"

Write-Log "CSV source  : $CsvFile" -Level INFO
Write-Log "XLSX cible  : $XlsxFile" -Level INFO

#endregion

# ==============================================================================
#region ETAPE 2 - VERIFICATION ET AUTO-INSTALLATION DE PYTHON + OPENPYXL
# ==============================================================================

Write-Log "Recherche de Python 3..." -Level INFO

$PythonExe   = $null
$PythonReady = $false

$PythonCandidates = [System.Collections.Generic.List[string]]@('python','python3','py')

foreach ($PyDir in @("$env:LOCALAPPDATA\Programs\Python", "C:\Python3*", "C:\tools\python*")) {
    try {
        Get-ChildItem $PyDir -Filter 'python.exe' -Recurse -ErrorAction SilentlyContinue |
            Sort-Object FullName -Descending | Select-Object -First 3 |
            ForEach-Object { $PythonCandidates.Add($_.FullName) }
    } catch {}
}

foreach ($Candidate in $PythonCandidates) {
    try {
        $TestPath = if (Test-Path $Candidate -ErrorAction SilentlyContinue) {
            $Candidate
        } else {
            $cmd = Get-Command $Candidate -ErrorAction SilentlyContinue
            if ($cmd) { $cmd.Source }
        }
        if ($TestPath -and (Test-Path $TestPath)) {
            $Ver = & $TestPath --version 2>&1
            if ("$Ver" -match 'Python 3\.([789]|1[0-9])') {
                $PythonExe = $TestPath
                Write-Log "Python trouve : $PythonExe ($Ver)" -Level SUCCESS
                break
            }
        }
    } catch { continue }
}

# Installation via winget si Python absent
if (-not $PythonExe) {
    Write-Log "Python 3 introuvable. Tentative d installation via winget..." -Level WARN
    try {
        $null = Get-Command winget -ErrorAction Stop
        $null = & winget install --id Python.Python.3.12 --silent `
            --accept-package-agreements --accept-source-agreements 2>&1
        $env:Path = [System.Environment]::GetEnvironmentVariable('Path','Machine') + ';' +
                    [System.Environment]::GetEnvironmentVariable('Path','User')
        $cmdPy = Get-Command python -ErrorAction SilentlyContinue
        $PythonExe = if ($cmdPy) { $cmdPy.Source } else { $null }
        if ($PythonExe) {
            Write-Log "Python installe via winget : $PythonExe" -Level SUCCESS
        } else {
            Write-Log "Installation winget terminee mais python toujours introuvable dans le PATH." -Level WARN
        }
    } catch {
        Write-Log "Installation Python via winget impossible : $_" -Level WARN
    }
}

if (-not $PythonExe) {
    Write-Log "Python 3 introuvable. Impossible de generer le fichier Excel." -Level ERROR
    Write-Log "Installez Python 3 depuis https://www.python.org/downloads/ puis relancez ce script." -Level INFO
    exit 1
}

# Verification / installation openpyxl
$ChkOut = & $PythonExe -c "import openpyxl; print(openpyxl.__version__)" 2>&1
if ($LASTEXITCODE -ne 0) {
    Write-Log "openpyxl absent, installation en cours..." -Level WARN
    $null = & $PythonExe -m pip install openpyxl --quiet 2>&1
    & $PythonExe -c "import openpyxl" 2>&1 | Out-Null
    if ($LASTEXITCODE -eq 0) {
        Write-Log "openpyxl installe avec succes." -Level SUCCESS
        $PythonReady = $true
    } else {
        Write-Log "Echec installation openpyxl. Verifiez votre connexion Internet." -Level ERROR
        exit 1
    }
} else {
    Write-Log "openpyxl $ChkOut OK" -Level SUCCESS
    $PythonReady = $true
}

#endregion

# ==============================================================================
#region ETAPE 3 - GENERATION DU FICHIER EXCEL
# ==============================================================================

Write-Log "Generation du fichier Excel en cours..." -Level INFO

# Extraction des metadonnees pour injection dans le script Python
$ElapsedMin    = if ($MetaLoaded -and $Meta.DureeMinutes) { $Meta.DureeMinutes } else { 0 }
$AllSrvCount   = if ($MetaLoaded -and $Meta.AllSrvCount)  { $Meta.AllSrvCount  } else { 0 }
$OnlineCount   = if ($MetaLoaded -and $Meta.OnlineCount)  { $Meta.OnlineCount  } else { 0 }
$OfflineCount  = if ($MetaLoaded -and $Meta.OfflineCount) { $Meta.OfflineCount } else { 0 }
$ErrorSrvCount = if ($MetaLoaded -and $Meta.ErrorSrvCount){ $Meta.ErrorSrvCount} else { 0 }
$Domaine       = if ($MetaLoaded -and $Meta.Domaine)      { $Meta.Domaine      } else { 'N/A' }
$DateAuditMeta = if ($MetaLoaded -and $Meta.DateAudit)    { $Meta.DateAudit    } else { (Get-Date -Format 'dd/MM/yyyy HH:mm') }

$OfflineSrvStr = if ($MetaLoaded -and $Meta.OfflineServers) { $Meta.OfflineServers -join '|' } else { '' }
$WinRMKOStr    = if ($MetaLoaded -and $Meta.WinRMKO)        { $Meta.WinRMKO -join '|'        } else { '' }
$ErrorSrvStr   = if ($MetaLoaded -and $Meta.ErrorServers)   { $Meta.ErrorServers -join ','    } else { '' }
$AuditErrStr   = if ($MetaLoaded -and $Meta.AuditErrors)    { ($Meta.AuditErrors | Select-Object -First 200) -join '||' } else { '' }

$PythonScript = @"
import csv, sys, os
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.chart import BarChart, Reference

csv_path  = r"$CsvFile"
xlsx_path = r"$XlsxFile"

# ── Palette couleurs ─────────────────────────────────────────
C_HEADER_BG = "1F4E79"; C_HEADER_FG = "FFFFFF"
C_ROW_ALT   = "EBF3FA"; C_ROW_NORM  = "FFFFFF"
C_BORDER    = "B8CCE4"
RISK_COLORS = {
    "CRITIQUE": ("C00000", "FFFFFF"),
    "ELEVE":    ("E74C3C", "FFFFFF"),
    "MOYEN":    ("FF9900", "FFFFFF"),
    "FAIBLE":   ("F0AD4E", "FFFFFF"),
}
C_WARN_BG = "FFF2CC"; C_WARN_FG = "7D6608"

def tb(color=C_BORDER):
    s = Side(border_style="thin", color=color)
    return Border(left=s, right=s, top=s, bottom=s)

def hdr(cell, bg=C_HEADER_BG, fg=C_HEADER_FG):
    cell.font      = Font(name="Arial", bold=True, color=fg, size=10)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    cell.border    = tb()

def data_cell(cell, bg=C_ROW_NORM, bold=False, fg="000000", center=False):
    cell.font      = Font(name="Arial", size=9, bold=bold, color=fg)
    cell.fill      = PatternFill("solid", fgColor=bg)
    cell.border    = tb()
    cell.alignment = Alignment(horizontal="center" if center else "left",
                               vertical="center", wrap_text=True)

# ── Lecture CSV ───────────────────────────────────────────────
with open(csv_path, encoding="utf-8-sig", newline="") as f:
    reader = csv.DictReader(f, delimiter=";")
    rows   = list(reader)
    hdrs   = list(rows[0].keys()) if rows else []

LABELS = {
    "Serveur": "Serveur", "NomPartage": "Nom du partage",
    "Description": "Description", "CheminLocal": "Chemin local",
    "ProprietaireNTFS": "Proprietaire NTFS", "NiveauRisque": "Niveau de risque",
    "DroitPartageEveryone": "Droit partage (Everyone)",
    "DroitNTFSEveryone": "Droit NTFS (Everyone)",
    "AutresControleTotal": "Autres Controle total",
    "TailleDossierMB": "Taille (Mo)", "LimitUtilisateurs": "Limite utilisateurs",
    "ErreurNTFS": "Erreur NTFS", "DateAudit": "Date audit",
}
WIDTHS = {
    "Serveur": 20, "NomPartage": 22, "Description": 28, "CheminLocal": 38,
    "ProprietaireNTFS": 28, "NiveauRisque": 14, "DroitPartageEveryone": 24,
    "DroitNTFSEveryone": 32, "AutresControleTotal": 36, "TailleDossierMB": 12,
    "LimitUtilisateurs": 14, "ErreurNTFS": 12, "DateAudit": 19,
}

def make_table_sheet(ws_t, sheet_rows, tbl_name):
    ws_t.sheet_view.showGridLines = False
    ws_t.append([LABELS.get(h, h) for h in hdrs])
    ws_t.row_dimensions[1].height = 34
    for cell in ws_t[1]:
        hdr(cell)
    risk_idx = (hdrs.index("NiveauRisque") + 1) if "NiveauRisque" in hdrs else None
    err_idx  = (hdrs.index("ErreurNTFS")  + 1) if "ErreurNTFS"  in hdrs else None
    for rn, row_d in enumerate(sheet_rows, start=2):
        ws_t.append([row_d.get(h, "") for h in hdrs])
        ws_t.row_dimensions[rn].height = 16
        bg_base = C_ROW_ALT if rn % 2 == 0 else C_ROW_NORM
        rv = (row_d.get("NiveauRisque") or "").upper()
        for cn, cell in enumerate(ws_t[rn], start=1):
            bg, fg, bold, center = bg_base, "000000", False, False
            if risk_idx and cn == risk_idx and rv in RISK_COLORS:
                bg, fg = RISK_COLORS[rv]; bold = center = True
            elif err_idx and cn == err_idx and str(cell.value or "").upper() == "OUI":
                bg, fg, bold = C_WARN_BG, C_WARN_FG, True
            data_cell(cell, bg=bg, bold=bold, fg=fg, center=center)
    for cn, cn_name in enumerate(hdrs, start=1):
        ws_t.column_dimensions[get_column_letter(cn)].width = WIDTHS.get(cn_name, 18)
    ws_t.freeze_panes = "A2"
    if sheet_rows:
        tbl = Table(displayName=tbl_name,
                    ref=f"A1:{get_column_letter(len(hdrs))}{len(sheet_rows)+1}")
        tbl.tableStyleInfo = TableStyleInfo(
            name="TableStyleMedium2", showRowStripes=True,
            showFirstColumn=False, showLastColumn=False, showColumnStripes=False)
        ws_t.add_table(tbl)

wb = Workbook()

# ── FEUILLE 1 : AUDIT COMPLET ─────────────────────────────────
ws = wb.active
ws.title = "Audit Partages"
make_table_sheet(ws, rows, "AuditPartages")

# ── FEUILLE 2 : RESUME EXECUTIF ───────────────────────────────
ws2 = wb.create_sheet("Resume")
ws2.sheet_view.showGridLines = False
ws2.column_dimensions["A"].width = 36
ws2.column_dimensions["B"].width = 20

ws2.merge_cells("A1:B1")
ws2["A1"] = "Resume de l audit de securite"
hdr(ws2["A1"])
ws2.row_dimensions[1].height = 34

risk_counts = {k: 0 for k in RISK_COLORS}
servers_exp = set()
for r in rows:
    rv = (r.get("NiveauRisque") or "").upper()
    if rv in risk_counts: risk_counts[rv] += 1
    if rv in RISK_COLORS: servers_exp.add(r.get("Serveur", ""))

meta = [
    ("Date de l audit",                   "$DateAuditMeta"),
    ("Domaine audite",                     "$Domaine"),
    ("Duree totale (minutes)",             $ElapsedMin),
    ("Serveurs dans l AD",                 $AllSrvCount),
    ("Serveurs joignables (WinRM OK)",     $OnlineCount),
    ("Serveurs hors ligne / WinRM KO",     $OfflineCount),
    ("Serveurs en erreur (audit)",         $ErrorSrvCount),
    ("Serveurs avec exposition detectee",  len(servers_exp)),
    ("Partages exposes (total)",           len([r for r in rows if (r.get("NiveauRisque") or "").upper() in RISK_COLORS])),
    (u"\u2500\u2500 Par niveau de risque \u2500\u2500", ""),
    ("   CRITIQUE",                        risk_counts["CRITIQUE"]),
    ("   ELEVE",                           risk_counts["ELEVE"]),
    ("   MOYEN",                           risk_counts["MOYEN"]),
    ("   FAIBLE",                          risk_counts["FAIBLE"]),
]

for rn, (lbl, val) in enumerate(meta, start=3):
    ws2[f"A{rn}"] = lbl
    ws2[f"B{rn}"] = val
    ws2.row_dimensions[rn].height = 20
    is_sep = str(lbl).startswith(u"\u2500")
    lbl_up = str(lbl).upper()
    for col in ["A","B"]:
        c = ws2[f"{col}{rn}"]
        if is_sep:
            c.font = Font(name="Arial", bold=True, size=9, color="1F4E79")
            c.fill = PatternFill("solid", fgColor="D6E4F0")
        else:
            c.font = Font(name="Arial", size=10, bold=(col == "B"))
            c.fill = PatternFill("solid", fgColor=(C_ROW_ALT if rn % 2 == 0 else C_ROW_NORM))
        c.border    = tb()
        c.alignment = Alignment(horizontal="center" if col == "B" else "left", vertical="center")
    for rk, (bg_c, fg_c) in RISK_COLORS.items():
        if rk in lbl_up:
            for col in ["A","B"]:
                ws2[f"{col}{rn}"].fill = PatternFill("solid", fgColor=bg_c)
                ws2[f"{col}{rn}"].font = Font(name="Arial", bold=True, size=10, color=fg_c)
            break

# Graphique en barres
if any(v > 0 for v in risk_counts.values()):
    ws2.column_dimensions["D"].width = 14
    ws2.column_dimensions["E"].width = 10
    ws2["D3"] = "Niveau"; ws2["E3"] = "Nb"
    hdr(ws2["D3"]); hdr(ws2["E3"])
    for idx, (lbl_c, val_c) in enumerate(
            [("CRITIQUE", risk_counts["CRITIQUE"]),
             ("ELEVE",    risk_counts["ELEVE"]),
             ("MOYEN",    risk_counts["MOYEN"]),
             ("FAIBLE",   risk_counts["FAIBLE"])], start=4):
        ws2[f"D{idx}"] = lbl_c; ws2[f"E{idx}"] = val_c
        for col in ["D","E"]:
            ws2[f"{col}{idx}"].border = tb()
            ws2[f"{col}{idx}"].font = Font(name="Arial", size=9)
    chart = BarChart()
    chart.type = "col"; chart.title = "Partages par niveau de risque"
    chart.y_axis.title = "Nombre"; chart.x_axis.title = "Niveau"
    chart.style = 10; chart.width = 14; chart.height = 10; chart.legend = None
    chart.add_data(Reference(ws2, min_col=5, min_row=3, max_row=7), titles_from_data=True)
    chart.set_categories(Reference(ws2, min_col=4, min_row=4, max_row=7))
    ws2.add_chart(chart, "D9")

# ── FEUILLES PAR NIVEAU DE RISQUE ─────────────────────────────
for level in ["CRITIQUE", "ELEVE", "MOYEN", "FAIBLE"]:
    level_rows = [r for r in rows if (r.get("NiveauRisque") or "").upper() == level]
    if not level_rows: continue
    ws_l = wb.create_sheet(f"Risque {level.capitalize()}")
    make_table_sheet(ws_l, level_rows, f"Risque_{level}")
    bg_l, fg_l = RISK_COLORS[level]
    for cell in ws_l[1]:
        hdr(cell, bg=bg_l, fg=fg_l)

# ── FEUILLE SERVEURS KO ────────────────────────────────────────
offline_list  = [x for x in """$OfflineSrvStr""".split("|") if x]
winrmko_list  = [x for x in """$WinRMKOStr""".split("|")   if x]
err_srv_list  = [x for x in "$ErrorSrvStr".split(",")       if x]

if offline_list or winrmko_list or err_srv_list:
    ws_ko = wb.create_sheet("Serveurs KO")
    ws_ko.sheet_view.showGridLines = False
    ws_ko.column_dimensions["A"].width = 28
    ws_ko.column_dimensions["B"].width = 36
    ws_ko.append(["Serveur", "Raison"])
    ws_ko.row_dimensions[1].height = 28
    for cell in ws_ko[1]: hdr(cell)
    rko = 2
    added = set()
    for srv in offline_list:
        if srv and srv not in added:
            ws_ko[f"A{rko}"] = srv; ws_ko[f"B{rko}"] = "Ping KO / Hors ligne"
            for c in ["A","B"]: ws_ko[f"{c}{rko}"].border = tb(); ws_ko[f"{c}{rko}"].font = Font(name="Arial",size=9)
            added.add(srv); rko += 1
    for srv in winrmko_list:
        if srv and srv not in added:
            ws_ko[f"A{rko}"] = srv; ws_ko[f"B{rko}"] = "Ping OK - WinRM KO (port 5985)"
            for c in ["A","B"]: ws_ko[f"{c}{rko}"].border = tb(); ws_ko[f"{c}{rko}"].font = Font(name="Arial",size=9)
            added.add(srv); rko += 1
    for srv in err_srv_list:
        if srv and srv not in added:
            ws_ko[f"A{rko}"] = srv; ws_ko[f"B{rko}"] = "WinRM timeout / Erreur audit"
            for c in ["A","B"]:
                ws_ko[f"{c}{rko}"].border = tb()
                ws_ko[f"{c}{rko}"].font = Font(name="Arial", size=9, color="C00000", bold=True)
            added.add(srv); rko += 1

# ── FEUILLE JOURNAL DES ERREURS D AUDIT ───────────────────────
audit_errs = [x for x in """$AuditErrStr""".split("||") if x]
if audit_errs:
    ws_er = wb.create_sheet("Erreurs Audit")
    ws_er.sheet_view.showGridLines = False
    ws_er.column_dimensions["A"].width = 90
    ws_er.append(["Detail des erreurs rencontrees pendant l audit"])
    hdr(ws_er["A1"])
    for idx, e in enumerate(audit_errs, start=2):
        ws_er[f"A{idx}"] = e
        ws_er[f"A{idx}"].font   = Font(name="Arial", size=8)
        ws_er[f"A{idx}"].border = tb()
        ws_er.row_dimensions[idx].height = 14

# ── ORDRE DES FEUILLES + FEUILLE ACTIVE = RESUME ─────────────
order = (["Resume", "Audit Partages"] +
         [s for s in wb.sheetnames if s.startswith("Risque ")] +
         [s for s in wb.sheetnames if s in ("Serveurs KO", "Erreurs Audit")])
for target_idx, sname in enumerate(order):
    if sname in wb.sheetnames:
        current_idx = wb.sheetnames.index(sname)
        if current_idx != target_idx:
            wb.move_sheet(sname, offset=target_idx - current_idx)

wb.active = wb["Resume"]
wb.save(xlsx_path)
print(f"OK:{xlsx_path}")
"@

$TempPy = [System.IO.Path]::GetTempFileName() -replace '\.tmp$', '.py'
try {
    [System.IO.File]::WriteAllText($TempPy, $PythonScript, [System.Text.Encoding]::UTF8)

    $PythonOutput = & $PythonExe $TempPy 2>&1
    $ExitCode     = $LASTEXITCODE

    if ($ExitCode -eq 0 -and ("$PythonOutput" -match '^OK:')) {
        Write-Log "Rapport Excel genere avec succes : $XlsxFile" -Level SUCCESS
    } else {
        Write-Log "Echec generation XLSX (code $ExitCode) : $PythonOutput" -Level ERROR
        exit 1
    }
} catch {
    Write-Log "Erreur execution Python : $_" -Level ERROR
    exit 1
} finally {
    if (Test-Path $TempPy -ErrorAction SilentlyContinue) {
        Remove-Item $TempPy -Force -ErrorAction SilentlyContinue
    }
}

#endregion

# ==============================================================================
#region RESUME FINAL
# ==============================================================================

$Duration = New-TimeSpan -Start $StartTime -End (Get-Date)

Write-Host @"

  +--------------------------------------------------------------+
  |                 MISE EN FORME TERMINEE                       |
  +--------------------------------------------------------------+
  | Duree           : $("{0:D2}m {1:D2}s" -f [int]$Duration.TotalMinutes, $Duration.Seconds)
  | Fichier Excel   : $XlsxFile
  +--------------------------------------------------------------+
"@ -ForegroundColor Green

# Ouvrir automatiquement le fichier Excel
try {
    $OpenExcel = Read-Host "`n  Voulez-vous ouvrir le fichier Excel maintenant ? (O/N)"
    if ($OpenExcel -match '^[Oo]') {
        Start-Process $XlsxFile
    }
} catch { }

#endregion
