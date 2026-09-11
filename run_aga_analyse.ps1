<#
.SYNOPSIS
    Kjører AGA-analysen (aga_analyse.py) fra prosjektmappen.

.DESCRIPTION
    1. Kontrollerer at Python er tilgjengelig.
    2. Oppretter et virtuelt miljø i .\.venv dersom det mangler.
    3. Aktiverer miljøet.
    4. Installerer pakkene i requirements.txt.
    5. Kjører aga_analyse.py.
    6. Viser hvor resultatfilen ble lagret.
    7. Stopper med en forståelig norsk feilmelding dersom noe går galt.
#>

$ErrorActionPreference = "Stop"

# Kjør alltid fra mappen dette skriptet ligger i, uansett hvor PowerShell ble startet.
$Prosjektmappe = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location -Path $Prosjektmappe

function Skriv-Feil {
    param([string]$Melding)
    Write-Host ""
    Write-Host "FEIL: $Melding" -ForegroundColor Red
    Write-Host ""
}

Write-Host "=== AGA-analyse ===" -ForegroundColor Cyan
Write-Host "Prosjektmappe: $Prosjektmappe"

# 1) Kontroller at Python er tilgjengelig
$PythonKommando = $null
foreach ($kandidat in @("python", "py")) {
    try {
        $versjon = & $kandidat --version 2>&1
        if ($LASTEXITCODE -eq 0) {
            $PythonKommando = $kandidat
            Write-Host "Fant Python via '$kandidat': $versjon"
            break
        }
    } catch {
        continue
    }
}

if (-not $PythonKommando) {
    Skriv-Feil "Fant ikke Python. Installer Python 3 fra https://www.python.org/downloads/ og sørg for at 'python' eller 'py' er tilgjengelig i PATH, kjør deretter dette skriptet på nytt."
    exit 1
}

# 2) Opprett virtuelt miljø dersom det mangler
$VenvMappe = Join-Path $Prosjektmappe ".venv"
$VenvPython = Join-Path $VenvMappe "Scripts\python.exe"

if (-not (Test-Path $VenvMappe)) {
    Write-Host "Oppretter virtuelt miljø i .\.venv ..."
    & $PythonKommando -m venv $VenvMappe
    if ($LASTEXITCODE -ne 0) {
        Skriv-Feil "Klarte ikke å opprette virtuelt miljø i .\.venv. Kontroller at Python-modulen 'venv' er tilgjengelig."
        exit 1
    }
}

if (-not (Test-Path $VenvPython)) {
    Skriv-Feil "Fant ikke $VenvPython etter oppretting av virtuelt miljø. Slett .\.venv-mappen og prøv på nytt."
    exit 1
}

# 3) "Aktiver" miljøet (vi kaller likevel python.exe direkte fra .venv for å unngå
#    problemer med PowerShell sin execution policy for Activate.ps1)
Write-Host "Bruker virtuelt miljø: $VenvPython"

# 4) Installer nødvendige pakker
$RequirementsFil = Join-Path $Prosjektmappe "requirements.txt"
if (-not (Test-Path $RequirementsFil)) {
    Skriv-Feil "Fant ikke requirements.txt i prosjektmappen ($Prosjektmappe)."
    exit 1
}

Write-Host "Installerer pakker fra requirements.txt ..."
& $VenvPython -m pip install --quiet --upgrade pip
& $VenvPython -m pip install --quiet -r $RequirementsFil
if ($LASTEXITCODE -ne 0) {
    Skriv-Feil "Klarte ikke å installere pakkene i requirements.txt. Kontroller internettforbindelsen og prøv på nytt."
    exit 1
}

# 5) Kjør analysen
Write-Host ""
Write-Host "Kjører aga_analyse.py ..." -ForegroundColor Cyan
& $VenvPython (Join-Path $Prosjektmappe "aga_analyse.py")
$AnalyseExitKode = $LASTEXITCODE

if ($AnalyseExitKode -ne 0) {
    Skriv-Feil "AGA-analysen feilet (avslutningskode $AnalyseExitKode). Se meldingen over og loggfilen i .\logs for detaljer."
    exit $AnalyseExitKode
}

# 6) Pek på filen de fleste vil åpne først (Python-utskriften over viser ALLE filene)
$OutputMappe = Join-Path $Prosjektmappe "output"
$FellesArbeidsbok = Get-ChildItem -Path $OutputMappe -Filter "AGA-analyse-*-alle-steg.xlsx" -ErrorAction SilentlyContinue |
    Sort-Object LastWriteTime -Descending | Select-Object -First 1

Write-Host ""
Write-Host "=== Fullført ===" -ForegroundColor Green
if ($FellesArbeidsbok) {
    Write-Host "Start her - felles arbeidsbok med alle steg: $($FellesArbeidsbok.FullName)"
} else {
    Write-Host "Analysen kjørte, men fant ingen AGA-analyse-*-alle-steg.xlsx i $OutputMappe - se skjermutskriften over."
}

$Oppslagsbehov = Join-Path $OutputMappe "AGA_oppslagsbehov.xlsx"
if (Test-Path $Oppslagsbehov) {
    Write-Host ""
    Write-Host "MERK: Offisiell AGA-kilde manglet ved denne kjøringen." -ForegroundColor Yellow
    Write-Host "Se: $Oppslagsbehov"
}
