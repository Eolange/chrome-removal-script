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

#region 5. Bloquer chrome.exe par ACL (icacls)
$chromePaths = @(
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe"
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if (-not $chromePaths) {
    Write-Log INFO "Chrome n'est pas installé sur ce poste."
} else {
    # Nettoyer les anciennes règles SRP qui peuvent bloquer tout le monde
    $srpCleanup = 'HKLM:\SOFTWARE\Policies\Microsoft\Windows\Safer\CodeIdentifiers\0\Paths'
    if (Test-Path $srpCleanup) {
        Get-ChildItem $srpCleanup -ErrorAction SilentlyContinue | ForEach-Object {
            $itemData = (Get-ItemProperty $_.PSPath -ErrorAction SilentlyContinue).ItemData
            if ($itemData -match 'chrome\.exe') {
                Remove-Item $_.PSPath -Recurse -Force -ErrorAction SilentlyContinue
                Write-Log SUCCESS "Ancienne règle SRP supprimée : $itemData"
            }
        }
    }

    foreach ($chromeExe in $chromePaths) {
        $chromeDir = Split-Path $chromeExe -Parent
        $chromeAppDir = Split-Path $chromeDir -Parent
        Write-Log INFO "Blocage de : $chromeExe"

        # 1. Reprendre possession de tout (répare les dégâts des exécutions précédentes)
        Write-Log INFO "  takeown sur $chromeAppDir..."
        & takeown /F $chromeAppDir /A /R /D O 2>&1 | Out-Null

        # 2. Réinitialiser TOUTES les ACL (héritage normal restauré)
        Write-Log INFO "  Reset ACL sur $chromeAppDir..."
        & icacls $chromeAppDir /reset /T /Q 2>&1 | Out-Null
        Write-Log SUCCESS "  Dossier restauré en permissions normales."

        # 3. Sur chaque .exe : bloquer l'exécution pour les non-admins
        #    Méthode : couper l'héritage + permissions explicites uniquement
        #    Admins et SYSTEM = FullControl, Users = rien du tout
        #    PAS de Deny (un Deny bloquerait aussi l'admin élevé car il est membre de Users)
        $exeFiles = Get-ChildItem -LiteralPath $chromeDir -Filter '*.exe' -ErrorAction SilentlyContinue
        foreach ($exe in $exeFiles) {
            $exePath = $exe.FullName
            Write-Log INFO "  Traitement : $($exe.Name)"

            # a) D'abord : supprimer tous les anciens Deny qui bloquent tout le monde
            $r = & icacls $exePath /remove:d '*S-1-5-32-545' 2>&1
            Write-Log INFO "    remove deny Users: $($r | Select-Object -First 1)"
            $r = & icacls $exePath /remove:d '*S-1-5-11' 2>&1
            Write-Log INFO "    remove deny AuthUsers: $($r | Select-Object -First 1)"

            # b) Donner FullControl explicite aux Admins et SYSTEM (AVANT de couper l'héritage)
            $r = & icacls $exePath /grant '*S-1-5-32-544:(F)' 2>&1
            Write-Log INFO "    grant Admins(F): $($r | Select-Object -First 1)"
            $r = & icacls $exePath /grant '*S-1-5-18:(F)' 2>&1
            Write-Log INFO "    grant SYSTEM(F): $($r | Select-Object -First 1)"

            # c) Couper l'héritage (les grants explicites ci-dessus survivent)
            $r = & icacls $exePath /inheritance:r 2>&1
            Write-Log INFO "    inheritance:r: $($r | Select-Object -First 1)"

            # d) Donner Read seul (PAS Execute) aux Users
            #    Nécessaire pour que Windows puisse lire le .exe (manifeste, prompt UAC)
            #    sinon même "Exécuter en tant qu'admin" échoue car Windows ne peut pas
            #    inspecter le fichier avant l'élévation.
            #    Read (R) ≠ ReadAndExecute (RX) : R ne permet PAS de lancer le programme.
            $r = & icacls $exePath /grant '*S-1-5-32-545:(R)' 2>&1
            Write-Log INFO "    grant Users(R): $($r | Select-Object -First 1)"

            # e) Vérifier
            $check = & icacls $exePath
            Write-Log INFO "    Permissions finales :"
            $check | ForEach-Object { if ($_.Trim()) { Write-Log INFO "      $_" } }
        }
    }
}
#endregion

Write-Log SUCCESS 'Traitement terminé.'
