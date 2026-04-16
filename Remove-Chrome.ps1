#requires -RunAsAdministrator

[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'

function Write-Log {
    param(
        [ValidateSet('INFO','SUCCESS','WARNING','ERROR')]
        [string]$Level,
        [string]$Message
    )
    Write-Host "[$Level] $Message"
}

function Test-IdentityResolvable {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Identity
    )
    try {
        $null = ([System.Security.Principal.NTAccount]$Identity).Translate([System.Security.Principal.SecurityIdentifier])
        $true
    }
    catch {
        $false
    }
}

function Remove-ItemSafe {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    try {
        if (Test-Path -LiteralPath $Path) {
            Remove-Item -LiteralPath $Path -Force -Recurse -ErrorAction Stop
            Write-Log 'SUCCESS' "Supprimé : $Path"
            return $true
        }
        else {
            Write-Log 'INFO' "Introuvable : $Path"
            return $false
        }
    }
    catch {
        Write-Log 'ERROR' "Impossible de supprimer $Path : $($_.Exception.Message)"
        return $false
    }
}

function Get-ActiveUserProfiles {
    $profiles = @()

    try {
        $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        $interactiveUser = $computerSystem.UserName

        Get-CimInstance Win32_UserProfile -ErrorAction SilentlyContinue | Where-Object {
            $_.Special -eq $false -and $_.LocalPath -like 'C:\Users\*'
        } | ForEach-Object {
            $profiles += [PSCustomObject]@{
                LocalPath = $_.LocalPath
                SID       = $_.SID
                Loaded    = $_.Loaded
            }
        }

        if ($interactiveUser) {
            $userNameOnly = $interactiveUser.Split('\')[-1]
            $profiles = $profiles | Sort-Object @{
                Expression = {
                    if ($_.LocalPath -match [regex]::Escape($userNameOnly)) { 0 }
                    elseif ($_.Loaded) { 1 }
                    else { 2 }
                }
            }
        }
    }
    catch {
        Write-Log 'WARNING' "Impossible d''énumérer les profils utilisateurs : $($_.Exception.Message)"
    }

    return $profiles
}

function Remove-ChromeDesktopShortcuts {
    $paths = @(
        "$env:PUBLIC\Desktop\Google Chrome.lnk",
        "$env:USERPROFILE\Desktop\Google Chrome.lnk"
    ) | Select-Object -Unique

    foreach ($path in $paths) {
        Remove-ItemSafe -Path $path | Out-Null
    }

    try {
        Get-ChildItem -Path 'C:\Users' -Directory -ErrorAction SilentlyContinue | ForEach-Object {
            $desktopShortcut = Join-Path $_.FullName 'Desktop\Google Chrome.lnk'
            Remove-ItemSafe -Path $desktopShortcut | Out-Null
        }
    }
    catch {
        Write-Log 'WARNING' "Impossible de parcourir C:\Users : $($_.Exception.Message)"
    }
}

function Disable-GoogleTasks {
    try {
        $tasks = Get-ScheduledTask -ErrorAction SilentlyContinue | Where-Object {
            $_.TaskName -match 'chrome|google'
        }

        if (-not $tasks) {
            Write-Log 'INFO' 'Aucune tâche planifiée liée à Chrome/Google trouvée.'
            return
        }

        foreach ($task in $tasks) {
            try {
                Disable-ScheduledTask -InputObject $task -ErrorAction SilentlyContinue | Out-Null
                Unregister-ScheduledTask -TaskName $task.TaskName -TaskPath $task.TaskPath -Confirm:$false -ErrorAction SilentlyContinue
                Write-Log 'SUCCESS' "Tâche supprimée : $($task.TaskPath)$($task.TaskName)"
            }
            catch {
                Write-Log 'ERROR' "Impossible de supprimer la tâche $($task.TaskPath)$($task.TaskName) : $($_.Exception.Message)"
            }
        }
    }
    catch {
        Write-Log 'ERROR' "Erreur lors de la gestion des tâches planifiées : $($_.Exception.Message)"
    }
}

function Stop-ChromeProcesses {
    try {
        Get-Process chrome -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Write-Log 'INFO' 'Processus Chrome arrêtés.'
    }
    catch {
        Write-Log 'WARNING' "Impossible d''arrêter Chrome : $($_.Exception.Message)"
    }
}

function Restart-ExplorerIfPossible {
    try {
        Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2
        Start-Process explorer.exe -ErrorAction SilentlyContinue
        Write-Log 'INFO' 'Explorateur redémarré.'
    }
    catch {
        Write-Log 'WARNING' "Impossible de redémarrer l''explorateur : $($_.Exception.Message)"
    }
}

function Get-ChromeExePath {
    $candidates = @(
        "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
        "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe"
    )

    foreach ($candidate in $candidates) {
        if ($candidate -and (Test-Path -LiteralPath $candidate)) {
            return $candidate
        }
    }

    return $null
}

function Test-IsChromeShortcut {
    param(
        [Parameter(Mandatory = $true)]
        [string]$LinkPath,
        [Parameter(Mandatory = $false)]
        $WshShell
    )

    try {
        $fileName = [System.IO.Path]::GetFileName($LinkPath)
        if ($fileName -match 'chrome') {
            return $true
        }

        if ($WshShell) {
            $shortcut = $WshShell.CreateShortcut($LinkPath)
            if ($shortcut.TargetPath -match '\\Google\\Chrome\\Application\\chrome\.exe$') {
                return $true
            }
        }
    }
    catch {
        Write-Log 'WARNING' "Impossible de lire le raccourci $LinkPath : $($_.Exception.Message)"
    }

    return $false
}

function Remove-ChromeTaskbarPins {
    $profiles = Get-ActiveUserProfiles
    $wshShell = $null

    try {
        $wshShell = New-Object -ComObject WScript.Shell
    }
    catch {
        Write-Log 'WARNING' "Impossible d''initialiser WScript.Shell : $($_.Exception.Message)"
    }

    foreach ($profile in $profiles) {
        $taskbarDir = Join-Path $profile.LocalPath 'AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar'

        try {
            if (-not (Test-Path -LiteralPath $taskbarDir)) {
                Write-Log 'INFO' "Dossier TaskBar introuvable : $taskbarDir"
                continue
            }

            $shell = New-Object -ComObject Shell.Application
            $folder = $shell.Namespace($taskbarDir)
            $links = Get-ChildItem -LiteralPath $taskbarDir -Filter '*.lnk' -ErrorAction SilentlyContinue

            foreach ($lnk in $links) {
                $isChrome = Test-IsChromeShortcut -LinkPath $lnk.FullName -WshShell $wshShell
                if (-not $isChrome) { continue }

                $item = $folder.ParseName($lnk.Name)
                $unpinned = $false

                if ($item) {
                    $unpinVerb = $item.Verbs() | Where-Object {
                        $_.Name -match 'Désépingler de la barre des tâches|Unpin from taskbar'
                    } | Select-Object -First 1

                    if ($unpinVerb) {
                        $unpinVerb.DoIt()
                        Start-Sleep -Milliseconds 500
                        $unpinned = $true
                        Write-Log 'SUCCESS' "Chrome désépinglé proprement : $($lnk.FullName)"
                    }
                }

                if (-not $unpinned) {
                    Remove-ItemSafe -Path $lnk.FullName | Out-Null
                    Write-Log 'WARNING' "Suppression directe du raccourci (fallback) : $($lnk.FullName)"
                }
            }
        }
        catch {
            Write-Log 'WARNING' "Impossible de nettoyer la taskbar pour $($profile.LocalPath) : $($_.Exception.Message)"
        }
    }

    $allUsersStartMenu = "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"
    Remove-ItemSafe -Path $allUsersStartMenu | Out-Null
}

function Test-ChromeTaskbarArtifact {
    $profiles = Get-ActiveUserProfiles
    $wshShell = $null

    try {
        $wshShell = New-Object -ComObject WScript.Shell
    }
    catch {}

    foreach ($profile in $profiles) {
        $taskbarDir = Join-Path $profile.LocalPath 'AppData\Roaming\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar'
        if (-not (Test-Path -LiteralPath $taskbarDir)) { continue }

        $links = Get-ChildItem -LiteralPath $taskbarDir -Filter '*.lnk' -ErrorAction SilentlyContinue
        foreach ($lnk in $links) {
            if (Test-IsChromeShortcut -LinkPath $lnk.FullName -WshShell $wshShell) {
                Write-Log 'WARNING' "Artefact taskbar Chrome encore présent : $($lnk.FullName)"
                return $true
            }
        }
    }

    return $false
}

function Reset-UserIconCache {
    param(
        [Parameter(Mandatory = $true)]
        [string]$UserProfilePath
    )

    $explorerCachePath = Join-Path $UserProfilePath 'AppData\Local\Microsoft\Windows\Explorer'
    $iconDb = Join-Path $UserProfilePath 'AppData\Local\IconCache.db'

    try {
        Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
        Start-Sleep -Seconds 2

        if (Test-Path -LiteralPath $explorerCachePath) {
            Get-ChildItem -LiteralPath $explorerCachePath -Filter 'iconcache*' -Force -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue

            Get-ChildItem -LiteralPath $explorerCachePath -Filter 'thumbcache*' -Force -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }

        if (Test-Path -LiteralPath $iconDb) {
            Remove-Item -LiteralPath $iconDb -Force -ErrorAction SilentlyContinue
        }

        Start-Process explorer.exe -ErrorAction SilentlyContinue
        Write-Log 'SUCCESS' "Cache d''icônes réinitialisé pour : $UserProfilePath"
    }
    catch {
        Write-Log 'WARNING' "Impossible de réinitialiser le cache d''icônes pour $UserProfilePath : $($_.Exception.Message)"
    }
}

function Invoke-ChromeTaskbarAutoRepair {
    $profiles = Get-ActiveUserProfiles

    Restart-ExplorerIfPossible
    Start-Sleep -Seconds 2

    if (Test-ChromeTaskbarArtifact) {
        Write-Log 'WARNING' 'Artefact taskbar Chrome détecté après nettoyage, reconstruction du cache d''icônes.'
        foreach ($profile in $profiles) {
            if ($profile.Loaded -or (Test-Path -LiteralPath $profile.LocalPath)) {
                Reset-UserIconCache -UserProfilePath $profile.LocalPath
            }
        }

        Start-Sleep -Seconds 2

        if (Test-ChromeTaskbarArtifact) {
            Write-Log 'WARNING' 'Un artefact taskbar Chrome semble encore présent après reconstruction du cache.'
        }
        else {
            Write-Log 'SUCCESS' 'Plus d''artefact taskbar Chrome détecté après reconstruction du cache.'
        }
    }
    else {
        Write-Log 'SUCCESS' 'Aucun artefact taskbar Chrome détecté après nettoyage.'
    }
}

function Set-ChromeRestrictionAcl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ChromeExePath
    )

    $chromeDir = Split-Path -Path $ChromeExePath -Parent
    $adminGroup = if (Test-IdentityResolvable -Identity 'Administrateurs') { 'Administrateurs' } else { 'Administrators' }

    Write-Log 'INFO' "Blocage ACL Chrome : $ChromeExePath"

    try {
        cmd /c "takeown /F `"$chromeDir`" /R /A" | Out-Null
        cmd /c "takeown /F `"$ChromeExePath`"" | Out-Null

        & icacls.exe "$chromeDir" /inheritance:r /T /C /Q | Out-Null
        & icacls.exe "$chromeDir" /grant:r "$adminGroup`:(OI)(CI)F" /grant:r "SYSTEM:(OI)(CI)F" /T /C /Q | Out-Null

        & icacls.exe "$ChromeExePath" /inheritance:r | Out-Null
        & icacls.exe "$ChromeExePath" /grant:r "$adminGroup`:F" /grant:r "SYSTEM:F" /C /Q | Out-Null

        foreach ($sid in @('*S-1-5-11', '*S-1-1-0')) {
            & icacls.exe "$ChromeExePath" /deny "${sid}:(RX)" /C /Q | Out-Null
        }

        Write-Log 'SUCCESS' "ACL Chrome appliquée : admins + SYSTEM conservés, utilisateurs standards bloqués"
    }
    catch {
        Write-Log 'ERROR' "Blocage ACL Chrome échoué : $($_.Exception.Message)"
    }
}

function Rename-ChromeExeFallback {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ChromeExePath
    )

    try {
        if (-not (Test-Path -LiteralPath $ChromeExePath)) {
            Write-Log 'INFO' "chrome.exe introuvable, renommage inutile."
            return
        }

        $disabledPath = "$ChromeExePath.disabled"

        if (Test-Path -LiteralPath $disabledPath) {
            Write-Log 'INFO' "Le binaire de secours existe déjà : $disabledPath"
            Remove-ItemSafe -Path $ChromeExePath | Out-Null
            return
        }

        Rename-Item -LiteralPath $ChromeExePath -NewName ([System.IO.Path]::GetFileName($disabledPath)) -Force -ErrorAction Stop
        Write-Log 'SUCCESS' "Renommage effectué : $ChromeExePath -> $disabledPath"
    }
    catch {
        Write-Log 'ERROR' "Impossible de renommer chrome.exe : $($_.Exception.Message)"
    }
}

Write-Log 'INFO' 'Début du traitement.'

Stop-ChromeProcesses
Remove-ChromeDesktopShortcuts
Disable-GoogleTasks
Remove-ChromeTaskbarPins
Invoke-ChromeTaskbarAutoRepair

$chromeExe = Get-ChromeExePath
if ($chromeExe) {
    Set-ChromeRestrictionAcl -ChromeExePath $chromeExe
}
else {
    Write-Log 'INFO' "Chrome n''est pas installé sur ce poste."
}

Restart-ExplorerIfPossible
Write-Log 'SUCCESS' 'Traitement terminé.'
