#requires -RunAsAdministrator
<#
.SYNOPSIS
    Bloque Google Chrome pour les utilisateurs non-administrateurs.
.DESCRIPTION
    - Arrête les processus Chrome
    - Supprime les raccourcis (bureau, taskbar, menu démarrer)
    - Désépingle Chrome de la barre des tâches (verbe Shell COM)
    - Désactive les tâches planifiées et services Google Update
    - Bloque chrome.exe par ACL (lecture seule pour Users) + SRP
    - Les admins conservent un accès complet
#>
[CmdletBinding()]
param()
$ErrorActionPreference = 'Continue'

#region Helpers
function Write-Log {
    param([ValidateSet('INFO','SUCCESS','WARNING','ERROR')][string]$Level, [string]$Message)
    $colors = @{ INFO='Cyan'; SUCCESS='Green'; WARNING='Yellow'; ERROR='Red' }
    Write-Host "[$Level] $Message" -ForegroundColor $colors[$Level]
}

function Get-LocalGroupName {
    param([string]$EnglishName, [string]$FrenchName)
    foreach ($name in $FrenchName, $EnglishName) {
        try {
            $null = ([System.Security.Principal.NTAccount]$name).Translate([System.Security.Principal.SecurityIdentifier])
            return $name
        } catch {}
    }
    return $EnglishName
}

$AdminGroup = Get-LocalGroupName -EnglishName 'Administrators' -FrenchName 'Administrateurs'
$UsersGroup = Get-LocalGroupName -EnglishName 'Users' -FrenchName 'Utilisateurs'
#endregion

#region 1. Arrêter Chrome
$procs = Get-Process chrome -ErrorAction SilentlyContinue
if ($procs) {
    $procs | Stop-Process -Force -ErrorAction SilentlyContinue
    Write-Log SUCCESS "$(@($procs).Count) processus Chrome arrêtés."
} else {
    Write-Log INFO 'Aucun processus Chrome en cours.'
}
#endregion

#region 2. Supprimer les raccourcis bureau + menu démarrer
$shortcutPaths = @(
    "$env:PUBLIC\Desktop\Google Chrome.lnk",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"
)
Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $shortcutPaths += Join-Path $_.FullName 'Desktop\Google Chrome.lnk'
}

foreach ($path in $shortcutPaths) {
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        Write-Log SUCCESS "Raccourci supprimé : $path"
    }
}
#endregion

#region 3. Désépingler Chrome de la barre des tâches
$wsh = New-Object -ComObject WScript.Shell -ErrorAction SilentlyContinue
$shell = New-Object -ComObject Shell.Application -ErrorAction SilentlyContinue

Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $qlRoot = Join-Path $_.FullName 'AppData\Roaming\Microsoft\Internet Explorer\Quick Launch'
    if (-not (Test-Path -LiteralPath $qlRoot)) { return }

    # Sous-dossiers à nettoyer
    $taskbarDir = Join-Path $qlRoot 'User Pinned\TaskBar'

    # a) Désépingler via verbe Shell COM (seule méthode fiable pour retirer l'icône)
    if ($shell -and (Test-Path -LiteralPath $taskbarDir)) {
        $folder = $shell.Namespace($taskbarDir)
        if ($folder) {
            foreach ($item in $folder.Items()) {
                $name = $item.Name
                if ($name -match 'chrome') {
                    foreach ($verb in $item.Verbs()) {
                        # FR: "Désépingler de la barre des tâches" / EN: "Unpin from taskbar"
                        if ($verb.Name -match 'pingler|Unpin|taskbar') {
                            $verb.DoIt()
                            Write-Log SUCCESS "Désépinglé de la taskbar : $name"
                            break
                        }
                    }
                }
            }
        }
    }

    # b) Supprimer les fichiers .lnk Chrome résiduels (taskbar + Quick Launch)
    Get-ChildItem -LiteralPath $qlRoot -Filter '*.lnk' -Recurse -Force -ErrorAction SilentlyContinue | ForEach-Object {
        $isChrome = $_.Name -match 'chrome'
        if (-not $isChrome -and $wsh) {
            try { $isChrome = $wsh.CreateShortcut($_.FullName).TargetPath -match 'chrome\.exe$' } catch {}
        }
        if ($isChrome) {
            Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
            Write-Log SUCCESS "Raccourci taskbar supprimé : $($_.FullName)"
        }
    }
}

# Redémarrer Explorer pour forcer le rafraîchissement de la barre des tâches
Get-Process explorer -ErrorAction SilentlyContinue | Stop-Process -Force -ErrorAction SilentlyContinue
Start-Sleep -Seconds 2
Start-Process explorer.exe -ErrorAction SilentlyContinue
Write-Log INFO 'Explorer redémarré (rafraîchissement taskbar).'
#endregion

#region 4. Désactiver tâches planifiées et services Google
Get-ScheduledTask -ErrorAction SilentlyContinue |
    Where-Object { $_.TaskName -match 'chrome|google' } |
    ForEach-Object {
        Unregister-ScheduledTask -TaskName $_.TaskName -TaskPath $_.TaskPath -Confirm:$false -ErrorAction SilentlyContinue
        Write-Log SUCCESS "Tâche supprimée : $($_.TaskPath)$($_.TaskName)"
    }

Get-Service -ErrorAction SilentlyContinue |
    Where-Object { $_.Name -match '^gupdate|^gupdatem|^GoogleChrome' } |
    ForEach-Object {
        Stop-Service -Name $_.Name -Force -ErrorAction SilentlyContinue
        Set-Service -Name $_.Name -StartupType Disabled -ErrorAction SilentlyContinue
        Write-Log SUCCESS "Service désactivé : $($_.Name)"
    }

foreach ($dir in "${env:ProgramFiles(x86)}\Google\Update", "${env:ProgramFiles}\Google\Update") {
    if (Test-Path -LiteralPath $dir) {
        Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log SUCCESS "Google Update supprimé : $dir"
    }
}
#endregion

#region 5. Bloquer chrome.exe par ACL (icacls) + SRP
$chromePaths = @(
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe"
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if (-not $chromePaths) {
    Write-Log INFO "Chrome n'est pas installé sur ce poste."
} else {
    foreach ($chromeExe in $chromePaths) {
        $chromeDir = Split-Path $chromeExe -Parent
        Write-Log INFO "Blocage de : $chromeExe"

        # D'abord : réparer le dossier parent si un ancien script a cassé ses ACL
        $chromeAppDir = Split-Path $chromeDir -Parent
        Write-Log INFO "Réparation ACL dossier : $chromeAppDir"
        & takeown /F $chromeAppDir /A /R /D O 2>$null | Out-Null
        & icacls $chromeAppDir /reset /T /Q 2>$null | Out-Null
        Write-Log SUCCESS "ACL dossier réinitialisé (héritage normal) : $chromeAppDir"

        # Bloquer chaque .exe : propriété admin + Users en Read seul (pas d'exécution)
        $exeFiles = Get-ChildItem -LiteralPath $chromeDir -Filter '*.exe' -ErrorAction SilentlyContinue
        foreach ($exe in $exeFiles) {
            $exePath = $exe.FullName
            # 1. Prendre possession pour Administrators
            & takeown /F $exePath /A 2>$null | Out-Null
            # 2. Couper l'héritage + supprimer toutes les permissions
            & icacls $exePath /inheritance:r 2>$null | Out-Null
            # 3. Donner FullControl aux admins et SYSTEM
            & icacls $exePath /grant "$($AdminGroup):(F)" 2>$null | Out-Null
            & icacls $exePath /grant "SYSTEM:(F)" 2>$null | Out-Null
            # 4. Users : Read seul (R = lire, pas RX = lire+exécuter)
            & icacls $exePath /grant "$($UsersGroup):(R)" 2>$null | Out-Null

            # Vérifier le résultat
            $check = & icacls $exePath 2>$null
            if ($check -match $AdminGroup) {
                Write-Log SUCCESS "ACL bloqué : $($exe.Name) ($AdminGroup=Full, $UsersGroup=Read)"
            } else {
                Write-Log ERROR "ACL possiblement échoué sur $($exe.Name)"
            }
        }
    }

    # SRP (Software Restriction Policy) - bloque l'exécution pour les non-admins
    $srpBase = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers'
    try {
        if (-not (Test-Path $srpBase)) { New-Item -Path $srpBase -Force | Out-Null }
        $current = Get-ItemProperty -Path $srpBase -ErrorAction SilentlyContinue
        if ($null -eq $current.DefaultLevel) { Set-ItemProperty $srpBase -Name 'DefaultLevel' -Value 262144 -Type DWord }
        Set-ItemProperty $srpBase -Name 'PolicyScope' -Value 1 -Type DWord

        $pathsBase = "$srpBase\0\Paths"
        if (-not (Test-Path $pathsBase)) { New-Item -Path $pathsBase -Force | Out-Null }

        $rules = @(
            @{ Guid='{B4A2D3A1-5C6D-4E8F-9A0B-1C2D3E4F5A6B}'; Path="${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe" },
            @{ Guid='{B4A2D3A1-5C6D-4E8F-9A0B-1C2D3E4F5A6C}'; Path="${env:ProgramFiles}\Google\Chrome\Application\chrome.exe" }
        )
        foreach ($rule in $rules) {
            if (-not $rule.Path) { continue }
            $rp = Join-Path $pathsBase $rule.Guid
            New-Item -Path $rp -Force | Out-Null
            Set-ItemProperty $rp -Name 'ItemData' -Value $rule.Path -Type String
            Set-ItemProperty $rp -Name 'SaferFlags' -Value 0 -Type DWord
            Write-Log SUCCESS "SRP bloqué : $($rule.Path)"
        }
    }
    catch {
        Write-Log ERROR "Configuration SRP échouée : $($_.Exception.Message)"
    }
}
#endregion

Write-Log SUCCESS 'Traitement terminé.'
