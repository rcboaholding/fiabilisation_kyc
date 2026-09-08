<#
  Redemarre proprement le serveur Django (aucune fenetre cmd creee), via waitress (WSGI prod).
  - arrete uniquement le processus enregistre dans django.pid (pas tous les python.exe)
  - relance en processus detache, fenetre masquee, sortie redirigee vers logs\django.log
#>
param(
    [string]$ProjectRoot = (Split-Path -Parent $PSScriptRoot),
    [string]$BindAddress = "10.170.82.20:8080"
)

$ErrorActionPreference = "Stop"
$LogDir  = Join-Path $ProjectRoot "logs"
New-Item -ItemType Directory -Force -Path $LogDir | Out-Null
$PidFile = Join-Path $ProjectRoot "django.pid"
$LogFile = Join-Path $LogDir "django.log"
$Python  = Join-Path $ProjectRoot "venv\Scripts\python.exe"

if (-not (Test-Path $Python)) { throw "Interpreteur introuvable : $Python" }

# --- 1. Arret cible du serveur precedent -------------------------------------
if (Test-Path $PidFile) {
    $oldPid = (Get-Content $PidFile -ErrorAction SilentlyContinue | Select-Object -First 1)
    if ($oldPid -match '^\d+$') {
        $proc = Get-Process -Id ([int]$oldPid) -ErrorAction SilentlyContinue
        if ($proc -and $proc.ProcessName -eq "python") {
            # runserver --noreload : un seul processus, pas d'enfant a traquer
            Stop-Process -Id ([int]$oldPid) -Force -ErrorAction SilentlyContinue
            Write-Output "Serveur Django precedent arrete (PID $oldPid)."
        }
    }
    Remove-Item $PidFile -Force -ErrorAction SilentlyContinue
}

# --- 2. Rotation simple du log (garde 7 jours) -------------------------------
if ((Test-Path $LogFile) -and ((Get-Item $LogFile).Length -gt 20MB)) {
    Move-Item $LogFile "$LogFile.1" -Force
}

# --- 3. Relance detachee, sans fenetre ---------------------------------------
$Host_, $Port_ = $BindAddress -split ':'
$proc = Start-Process -FilePath $Python `
    -ArgumentList @("-m", "waitress", "--host=$Host_", "--port=$Port_", "Fiabilisation_kyc.wsgi:application") `
    -WorkingDirectory $ProjectRoot `
    -WindowStyle Hidden `
    -RedirectStandardOutput $LogFile `
    -RedirectStandardError "$LogFile.err" `
    -PassThru

Start-Sleep -Seconds 3
if ($proc.HasExited) {
    throw "Le serveur Django s'est arrete immediatement (code $($proc.ExitCode)). Voir $LogFile.err"
}
$proc.Id | Set-Content -Path $PidFile -Encoding ascii
Write-Output "Serveur Django lance sur $BindAddress (PID $($proc.Id))."
exit 0
