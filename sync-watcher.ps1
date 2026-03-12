#!/usr/bin/env pwsh
<#
.SYNOPSIS
    Surveille le dossier OneDrive/SharePoint synchronisé et push automatiquement
    les plannings vers GitHub.

.DESCRIPTION
    Ce script tourne en arrière-plan et détecte les modifications de fichiers
    "Plannings YYYY SXX*.xlsx" dans le dossier OneDrive synchronisé.
    Quand un fichier change, il le copie dans le repo Git local,
    commit et push automatiquement.

.USAGE
    .\sync-watcher.ps1                          # Lancement interactif (demande les chemins)
    .\sync-watcher.ps1 -OneDrivePath "C:\..." -RepoPath "C:\..."  # Chemins explicites
    .\sync-watcher.ps1 -Install                 # Installe en tâche planifiée (démarrage auto)

.NOTES
    Nécessite : Git installé et configuré (git push doit fonctionner sans mot de passe)
#>

param(
    [string]$OneDrivePath = "",
    [string]$RepoPath = "",
    [switch]$Install,
    [switch]$Uninstall,
    [int]$Debounce = 10  # Secondes d'attente après la dernière modification
)

# ── Configuration ────────────────────────────────────────────────────────────

$PLANNING_PATTERN = "Plannings *.xlsx"
$TASK_NAME = "UrbanSoccer-PlanningSync"
$LOG_FILE = Join-Path $PSScriptRoot "sync-watcher.log"

# ── Fonctions utilitaires ────────────────────────────────────────────────────

function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] [$Level] $Message"
    Write-Host $line
    Add-Content -Path $LOG_FILE -Value $line -ErrorAction SilentlyContinue
}

function Find-OneDrivePlanningFolder {
    <# Cherche automatiquement le dossier Plannings dans OneDrive/SharePoint sync #>
    $searchPaths = @(
        "$env:USERPROFILE\OneDrive*\02. USP*\General\RH\Plannings",
        "$env:USERPROFILE\OneDrive*\USP*\General\RH\Plannings",
        "$env:USERPROFILE\*compagnie des alpes*\02. USP*\General\RH\Plannings",
        "$env:USERPROFILE\*compagnie des alpes*\USP*\General\RH\Plannings",
        "$env:USERPROFILE\OneDrive*\*\02. USP*\General\RH\Plannings"
    )

    foreach ($pattern in $searchPaths) {
        $found = Get-Item -Path $pattern -ErrorAction SilentlyContinue | Select-Object -First 1
        if ($found) {
            return $found.FullName
        }
    }

    # Recherche plus large
    $oneDriveFolders = Get-ChildItem "$env:USERPROFILE" -Directory | Where-Object {
        $_.Name -like "OneDrive*" -or $_.Name -like "*compagnie*" -or $_.Name -like "*SharePoint*"
    }

    foreach ($od in $oneDriveFolders) {
        $found = Get-ChildItem -Path $od.FullName -Recurse -Directory -Filter "Plannings" -ErrorAction SilentlyContinue |
            Where-Object { $_.FullName -like "*RH*" } | Select-Object -First 1
        if ($found) {
            return $found.FullName
        }
    }

    return $null
}

function Find-GitRepo {
    <# Cherche le repo planning-urbansoccer sur le disque #>
    $searchPaths = @(
        "$env:USERPROFILE\Documents\planning-urbansoccer",
        "$env:USERPROFILE\Desktop\planning-urbansoccer",
        "$env:USERPROFILE\Source\Repos\planning-urbansoccer",
        "$env:USERPROFILE\repos\planning-urbansoccer",
        "$env:USERPROFILE\git\planning-urbansoccer",
        "C:\Users\$env:USERNAME\planning-urbansoccer"
    )

    foreach ($path in $searchPaths) {
        if (Test-Path (Join-Path $path ".git")) {
            return $path
        }
    }

    # Recherche par git
    $gitResult = & git -C "$env:USERPROFILE" rev-parse --show-toplevel 2>$null
    if ($gitResult -and (Split-Path $gitResult -Leaf) -eq "planning-urbansoccer") {
        return $gitResult
    }

    return $null
}

function Install-GitRepo {
    <# Clone le repo automatiquement si absent #>
    $defaultPath = Join-Path $env:USERPROFILE "planning-urbansoccer"
    Write-Host ""
    Write-Host "  Le repo n'a pas ete trouve. Clonage automatique..." -ForegroundColor Yellow
    & git clone "https://github.com/OhLaPey/planning-urbansoccer.git" $defaultPath 2>&1
    if ($LASTEXITCODE -eq 0 -and (Test-Path (Join-Path $defaultPath ".git"))) {
        Write-Host "  Clone dans : $defaultPath" -ForegroundColor Green
        return $defaultPath
    }
    Write-Host "  Echec du clonage." -ForegroundColor Red
    return $null
}

function Test-IsTemporaryFile {
    param([string]$FileName)
    return ($FileName.StartsWith("~`$") -or $FileName.StartsWith(".~") -or $FileName.EndsWith(".tmp"))
}

function Test-IsPlanningFile {
    param([string]$FileName)
    return ($FileName -match "^Plannings\s+\d{4}\s+S\d{1,2}.*\.xlsx$")
}

function Sync-FileToGitHub {
    param(
        [string]$SourceFile,
        [string]$RepoPath
    )

    $fileName = Split-Path $SourceFile -Leaf
    $destFile = Join-Path $RepoPath $fileName

    Write-Log "Copie : $fileName -> repo"
    Copy-Item -Path $SourceFile -Destination $destFile -Force

    # Git add + commit + push
    Push-Location $RepoPath
    try {
        & git add $fileName 2>&1 | Out-Null

        # Vérifier s'il y a vraiment des changements
        $status = & git status --porcelain $fileName 2>&1
        if (-not $status) {
            Write-Log "Aucun changement détecté pour $fileName (fichier identique)" "INFO"
            return $false
        }

        # Extraire le numéro de semaine pour le message
        $weekMatch = [regex]::Match($fileName, "S(\d{1,2})")
        $weekInfo = if ($weekMatch.Success) { "S$($weekMatch.Groups[1].Value)" } else { $fileName }

        $commitMsg = "Sync planning SharePoint : $weekInfo"
        Write-Log "Commit : $commitMsg"
        & git commit -m $commitMsg 2>&1 | Out-Null

        Write-Log "Push vers GitHub..."
        $pushResult = & git push 2>&1
        if ($LASTEXITCODE -ne 0) {
            Write-Log "Erreur push : $pushResult" "ERROR"
            # Retry avec backoff
            foreach ($wait in @(2, 4, 8, 16)) {
                Write-Log "Retry dans ${wait}s..." "WARN"
                Start-Sleep -Seconds $wait
                $pushResult = & git push 2>&1
                if ($LASTEXITCODE -eq 0) { break }
            }
        }

        if ($LASTEXITCODE -eq 0) {
            Write-Log "Push OK pour $fileName" "INFO"
            return $true
        } else {
            Write-Log "Push échoué après 4 tentatives : $pushResult" "ERROR"
            return $false
        }
    }
    finally {
        Pop-Location
    }
}

# ── Installation en tâche planifiée ──────────────────────────────────────────

function Install-AsScheduledTask {
    param([string]$OneDrivePath, [string]$RepoPath)

    $scriptPath = $PSCommandPath
    $action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument @"
-WindowStyle Hidden -ExecutionPolicy Bypass -File "$scriptPath" -OneDrivePath "$OneDrivePath" -RepoPath "$RepoPath"
"@

    $trigger = New-ScheduledTaskTrigger -AtLogOn
    $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)
    $principal = New-ScheduledTaskPrincipal -UserId $env:USERNAME -LogonType Interactive -RunLevel Limited

    Register-ScheduledTask -TaskName $TASK_NAME -Action $action -Trigger $trigger -Settings $settings -Principal $principal -Force

    Write-Log "Tache planifiee '$TASK_NAME' installee (demarrage automatique a la connexion)" "INFO"
    Write-Host ""
    Write-Host "Le script demarrera automatiquement a chaque connexion Windows." -ForegroundColor Green
    Write-Host "Pour le desinstaller : .\sync-watcher.ps1 -Uninstall" -ForegroundColor Yellow
}

function Uninstall-ScheduledTask {
    Unregister-ScheduledTask -TaskName $TASK_NAME -Confirm:$false -ErrorAction SilentlyContinue
    Write-Log "Tache planifiee '$TASK_NAME' supprimee" "INFO"
    Write-Host "Tache planifiee supprimee." -ForegroundColor Green
}

# ── Boucle principale ────────────────────────────────────────────────────────

function Start-Watcher {
    param([string]$OneDrivePath, [string]$RepoPath)

    Write-Host ""
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host "  UrbanSoccer - Sync Planning SharePoint -> GitHub" -ForegroundColor Cyan
    Write-Host "============================================================" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "  Dossier surveille : $OneDrivePath" -ForegroundColor White
    Write-Host "  Repo Git local    : $RepoPath" -ForegroundColor White
    Write-Host "  Log               : $LOG_FILE" -ForegroundColor White
    Write-Host ""
    Write-Host "  En attente de modifications..." -ForegroundColor Green
    Write-Host "  (Ctrl+C pour arreter)" -ForegroundColor DarkGray
    Write-Host ""

    Write-Log "Demarrage watcher : $OneDrivePath -> $RepoPath"

    # Créer le FileSystemWatcher
    $watcher = New-Object System.IO.FileSystemWatcher
    $watcher.Path = $OneDrivePath
    $watcher.Filter = $PLANNING_PATTERN
    $watcher.NotifyFilter = [System.IO.NotifyFilters]::FileName -bor
                            [System.IO.NotifyFilters]::LastWrite -bor
                            [System.IO.NotifyFilters]::Size
    $watcher.IncludeSubdirectories = $false
    $watcher.EnableRaisingEvents = $true

    # Table pour le debounce (éviter les syncs multiples pour une même sauvegarde)
    $pendingFiles = @{}
    $lastSync = @{}

    try {
        while ($true) {
            # Attendre un événement (timeout 2 secondes pour vérifier le debounce)
            $result = $watcher.WaitForChanged(
                [System.IO.WatcherChangeTypes]::Changed -bor
                [System.IO.WatcherChangeTypes]::Created,
                2000
            )

            if (-not $result.TimedOut) {
                $fileName = $result.Name
                $filePath = Join-Path $OneDrivePath $fileName

                # Ignorer les fichiers temporaires
                if (Test-IsTemporaryFile $fileName) { continue }
                if (-not (Test-IsPlanningFile $fileName)) { continue }

                Write-Log "Modification detectee : $fileName"
                $pendingFiles[$fileName] = Get-Date
            }

            # Traiter les fichiers en attente (après le debounce)
            $now = Get-Date
            $toProcess = @()
            $toRemove = @()

            foreach ($entry in $pendingFiles.GetEnumerator()) {
                $elapsed = ($now - $entry.Value).TotalSeconds
                if ($elapsed -ge $Debounce) {
                    $toProcess += $entry.Key
                    $toRemove += $entry.Key
                }
            }

            foreach ($key in $toRemove) {
                $pendingFiles.Remove($key)
            }

            foreach ($fileName in $toProcess) {
                $filePath = Join-Path $OneDrivePath $fileName

                # Vérifier que le fichier existe et n'est pas verrouillé
                if (-not (Test-Path $filePath)) { continue }

                # Attendre que le fichier ne soit plus verrouillé (Excel peut le bloquer)
                $retries = 0
                $unlocked = $false
                while ($retries -lt 5) {
                    try {
                        $stream = [System.IO.File]::Open($filePath, 'Open', 'Read', 'ReadWrite')
                        $stream.Close()
                        $unlocked = $true
                        break
                    }
                    catch {
                        $retries++
                        Write-Log "Fichier verrouille, retry $retries/5..." "WARN"
                        Start-Sleep -Seconds 3
                    }
                }

                if (-not $unlocked) {
                    Write-Log "Impossible d'acceder a $fileName (verrouille)" "ERROR"
                    continue
                }

                # Sync
                $success = Sync-FileToGitHub -SourceFile $filePath -RepoPath $RepoPath
                if ($success) {
                    $lastSync[$fileName] = $now
                    Write-Host "  -> $fileName synchronise avec succes" -ForegroundColor Green
                }
            }
        }
    }
    finally {
        $watcher.Dispose()
        Write-Log "Watcher arrete"
    }
}

# ── Point d'entrée ───────────────────────────────────────────────────────────

# Désinstallation
if ($Uninstall) {
    Uninstall-ScheduledTask
    exit 0
}

# Résolution des chemins
if (-not $OneDrivePath) {
    Write-Host ""
    Write-Host "Recherche du dossier Plannings OneDrive/SharePoint..." -ForegroundColor Yellow
    $OneDrivePath = Find-OneDrivePlanningFolder

    if ($OneDrivePath) {
        Write-Host "  Trouve : $OneDrivePath" -ForegroundColor Green
        $confirm = Read-Host "  Utiliser ce dossier ? (O/n)"
        if ($confirm -eq "n" -or $confirm -eq "N") {
            $OneDrivePath = Read-Host "  Entrez le chemin du dossier Plannings"
        }
    }
    else {
        Write-Host "  Dossier non trouve automatiquement." -ForegroundColor Red
        Write-Host ""
        Write-Host "  Assurez-vous d'avoir clique sur 'Ajouter un raccourci vers OneDrive'" -ForegroundColor Yellow
        Write-Host "  dans SharePoint, puis synchronise le dossier." -ForegroundColor Yellow
        Write-Host ""
        $OneDrivePath = Read-Host "  Entrez le chemin du dossier Plannings synchronise"
    }
}

if (-not $RepoPath) {
    Write-Host ""
    Write-Host "Recherche du repo planning-urbansoccer..." -ForegroundColor Yellow
    $RepoPath = Find-GitRepo

    if ($RepoPath) {
        Write-Host "  Trouve : $RepoPath" -ForegroundColor Green
        $confirm = Read-Host "  Utiliser ce repo ? (O/n)"
        if ($confirm -eq "n" -or $confirm -eq "N") {
            $RepoPath = Read-Host "  Entrez le chemin du repo Git"
        }
    }
    else {
        Write-Host "  Repo non trouve. Clonage automatique..." -ForegroundColor Yellow
        $RepoPath = Install-GitRepo
        if (-not $RepoPath) {
            $RepoPath = Read-Host "  Entrez le chemin du repo Git manuellement"
        }
    }
}

# Validation
if (-not (Test-Path $OneDrivePath)) {
    Write-Host "ERREUR : Le dossier '$OneDrivePath' n'existe pas." -ForegroundColor Red
    exit 1
}

if (-not (Test-Path (Join-Path $RepoPath ".git"))) {
    Write-Host "ERREUR : '$RepoPath' n'est pas un repo Git." -ForegroundColor Red
    exit 1
}

# Installation en tâche planifiée
if ($Install) {
    Install-AsScheduledTask -OneDrivePath $OneDrivePath -RepoPath $RepoPath
    # Lancer aussi immédiatement
    Start-Watcher -OneDrivePath $OneDrivePath -RepoPath $RepoPath
}
else {
    Start-Watcher -OneDrivePath $OneDrivePath -RepoPath $RepoPath
}
