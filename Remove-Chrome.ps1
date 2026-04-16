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
        }
        else {
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
    }
    catch {
        Write-Log 'ERROR' "Erreur lors de la gestion des tâches planifiées : $($_.Exception.Message)"
    }
    $googleServices = Get-Service -ErrorAction SilentlyContinue | Where-Object {
        $_.Name -match '^gupdate|^gupdatem|^GoogleChrome'
    }
    if (-not $googleServices) {
        Write-Log 'INFO' 'Aucun service Google Update trouvé.'
    }
    else {
        foreach ($svc in $googleServices) {
            try {
                Stop-Service -Name $svc.Name -Force -ErrorAction SilentlyContinue
                Set-Service -Name $svc.Name -StartupType Disabled -ErrorAction Stop
                Write-Log 'SUCCESS' "Service désactivé : $($svc.Name) ($($svc.DisplayName))"
            }
            catch {
                Write-Log 'ERROR' "Impossible de désactiver le service $($svc.Name) : $($_.Exception.Message)"
            }
        }
    }
    $googleUpdatePaths = @(
        "${env:ProgramFiles(x86)}\Google\Update",
        "${env:ProgramFiles}\Google\Update"
    )
    foreach ($updateDir in $googleUpdatePaths) {
        if (Test-Path -LiteralPath $updateDir) {
            Remove-ItemSafe -Path $updateDir | Out-Null
        }
        else {
            Write-Log 'INFO' "Dossier Google Update introuvable : $updateDir"
        }
    }
}

function Stop-ChromeProcesses {
    try {
        $procs = Get-Process chrome -ErrorAction SilentlyContinue
        if ($procs) {
            $count = @($procs).Count
            $procs | Stop-Process -Force -ErrorAction SilentlyContinue
            Write-Log 'SUCCESS' "$count processus Chrome arrêtés."
        }
        else {
            Write-Log 'INFO' 'Aucun processus Chrome en cours.'
        }
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
        Write-Log 'INFO' "Traitement taskbar pour : $($profile.LocalPath)"
        $quickLaunchRoot = Join-Path $profile.LocalPath 'AppData\Roaming\Microsoft\Internet Explorer\Quick Launch'
        # Liste des dossiers à nettoyer
        $dirsToClean = @(
            (Join-Path $quickLaunchRoot 'User Pinned\TaskBar'),
            $quickLaunchRoot
        )
        # Ajouter les sous-dossiers de ImplicitAppShortcuts
        $implicitDir = Join-Path $quickLaunchRoot 'User Pinned\ImplicitAppShortcuts'
        if (Test-Path -LiteralPath $implicitDir) {
            Get-ChildItem -LiteralPath $implicitDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $dirsToClean += $_.FullName
            }
        }
        foreach ($dir in $dirsToClean) {
            try {
                if (-not (Test-Path -LiteralPath $dir)) { continue }
                $links = Get-ChildItem -LiteralPath $dir -Filter '*.lnk' -ErrorAction SilentlyContinue
                foreach ($lnk in $links) {
                    $isChrome = Test-IsChromeShortcut -LinkPath $lnk.FullName -WshShell $wshShell
                    if (-not $isChrome) { continue }
                    Remove-ItemSafe -Path $lnk.FullName | Out-Null
                    Write-Log 'SUCCESS' "Raccourci Chrome supprimé : $($lnk.FullName)"
                }
                # Supprimer les dossiers ImplicitAppShortcuts vides après nettoyage
                if ($dir -match 'ImplicitAppShortcuts\\' -and (Test-Path -LiteralPath $dir)) {
                    $remaining = Get-ChildItem -LiteralPath $dir -ErrorAction SilentlyContinue
                    if (-not $remaining) {
                        Remove-Item -LiteralPath $dir -Force -ErrorAction SilentlyContinue
                        Write-Log 'INFO' "Dossier ImplicitAppShortcuts vide supprimé : $dir"
                    }
                }
            }
            catch {
                Write-Log 'WARNING' "Impossible de nettoyer $dir : $($_.Exception.Message)"
            }
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
        if (Test-Path -LiteralPath $explorerCachePath) {
            Get-ChildItem -LiteralPath $explorerCachePath -Filter 'iconcache*' -Force -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
            Get-ChildItem -LiteralPath $explorerCachePath -Filter 'thumbcache*' -Force -ErrorAction SilentlyContinue |
                Remove-Item -Force -ErrorAction SilentlyContinue
        }
        if (Test-Path -LiteralPath $iconDb) {
            Remove-Item -LiteralPath $iconDb -Force -ErrorAction SilentlyContinue
        }
        Write-Log 'SUCCESS' "Cache d''icônes nettoyé pour : $UserProfilePath"
    }
    catch {
        Write-Log 'WARNING' "Impossible de réinitialiser le cache d''icônes pour $UserProfilePath : $($_.Exception.Message)"
    }
}

function Invoke-ChromeTaskbarAutoRepair {
    # Nettoyage complémentaire : supprimer tout raccourci Chrome résiduel dans les profils
    $profiles = Get-ActiveUserProfiles
    $wshShell = $null
    try { $wshShell = New-Object -ComObject WScript.Shell } catch {}
    $totalRemoved = 0
    foreach ($profile in $profiles) {
        $quickLaunchRoot = Join-Path $profile.LocalPath 'AppData\Roaming\Microsoft\Internet Explorer\Quick Launch'
        if (-not (Test-Path -LiteralPath $quickLaunchRoot)) { continue }
        # Rechercher récursivement tous les .lnk Chrome dans Quick Launch
        try {
            $allLinks = Get-ChildItem -LiteralPath $quickLaunchRoot -Filter '*.lnk' -Recurse -Force -ErrorAction SilentlyContinue
            foreach ($lnk in $allLinks) {
                if (Test-IsChromeShortcut -LinkPath $lnk.FullName -WshShell $wshShell) {
                    Remove-Item -LiteralPath $lnk.FullName -Force -ErrorAction SilentlyContinue
                    Write-Log 'SUCCESS' "Raccourci Chrome résiduel supprimé : $($lnk.FullName)"
                    $totalRemoved++
                }
            }
        }
        catch {
            Write-Log 'WARNING' "Impossible de parcourir $quickLaunchRoot : $($_.Exception.Message)"
        }
        # Nettoyer les dossiers ImplicitAppShortcuts vides
        $implicitDir = Join-Path $quickLaunchRoot 'User Pinned\ImplicitAppShortcuts'
        if (Test-Path -LiteralPath $implicitDir) {
            Get-ChildItem -LiteralPath $implicitDir -Directory -ErrorAction SilentlyContinue | ForEach-Object {
                $remaining = Get-ChildItem -LiteralPath $_.FullName -ErrorAction SilentlyContinue
                if (-not $remaining) {
                    Remove-Item -LiteralPath $_.FullName -Force -ErrorAction SilentlyContinue
                    Write-Log 'INFO' "Dossier vide supprimé : $($_.FullName)"
                }
            }
        }
    }
    if ($totalRemoved -eq 0) {
        Write-Log 'INFO' 'Aucun raccourci Chrome résiduel trouvé.'
    }
    else {
        Write-Log 'SUCCESS' "$totalRemoved raccourci(s) Chrome résiduel(s) supprimé(s)."
    }
}

function Set-ChromeRestrictionAcl {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ChromeExePath
    )
    $chromeDir = Split-Path -Path $ChromeExePath -Parent
    $chromeAppDir = Split-Path -Path $chromeDir -Parent
    $adminGroup = if (Test-IdentityResolvable -Identity 'Administrateurs') { 'Administrateurs' } else { 'Administrators' }
    Write-Log 'INFO' "Blocage ACL Chrome : $chromeAppDir"
    Write-Log 'INFO' "Groupe admin détecté : $adminGroup"
    # Activer le privilège SeTakeOwnershipPrivilege pour pouvoir changer le propriétaire
    $tokenPriv = @'
using System;
using System.Runtime.InteropServices;
public class TokenPriv {
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool AdjustTokenPrivileges(IntPtr h, bool d, ref TOKEN_PRIVILEGES n, int l, IntPtr p, IntPtr r);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool OpenProcessToken(IntPtr h, uint a, out IntPtr t);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool LookupPrivilegeValue(string s, string n, out long l);
    struct TOKEN_PRIVILEGES { public int Count; public long Luid; public int Attr; }
    public static void Enable(string priv) {
        IntPtr token; long luid;
        OpenProcessToken((IntPtr)(-1), 0x28, out token);
        LookupPrivilegeValue(null, priv, out luid);
        TOKEN_PRIVILEGES tp = new TOKEN_PRIVILEGES { Count = 1, Luid = luid, Attr = 2 };
        AdjustTokenPrivileges(token, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
    }
}
'@
    try { Add-Type $tokenPriv -ErrorAction SilentlyContinue } catch {}
    try {
        [TokenPriv]::Enable('SeTakeOwnershipPrivilege')
        [TokenPriv]::Enable('SeRestorePrivilege')
        Write-Log 'INFO' 'Privilèges SeTakeOwnership et SeRestore activés.'
    }
    catch {
        Write-Log 'WARNING' "Impossible d''activer les privilèges : $($_.Exception.Message)"
    }
    try {
        $adminAccount = New-Object System.Security.Principal.NTAccount($adminGroup)
        # Prendre possession du dossier racine
        $acl = Get-Acl -LiteralPath $chromeAppDir
        $acl.SetOwner($adminAccount)
        Set-Acl -LiteralPath $chromeAppDir -AclObject $acl -ErrorAction Stop
        Write-Log 'SUCCESS' "Propriétaire changé sur $chromeAppDir"
        # Appliquer les ACL sur le dossier racine
        $acl = Get-Acl -LiteralPath $chromeAppDir
        $acl.SetAccessRuleProtection($true, $false)
        $acl.Access | ForEach-Object { $acl.RemoveAccessRule($_) | Out-Null }
        $adminRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $adminGroup, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            'SYSTEM', 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $acl.AddAccessRule($adminRule)
        $acl.AddAccessRule($systemRule)
        Set-Acl -LiteralPath $chromeAppDir -AclObject $acl -ErrorAction Stop
        Write-Log 'SUCCESS' "ACL appliquée sur $chromeAppDir (racine)"
        # Appliquer récursivement
        $errors = 0
        $items = @(Get-ChildItem -LiteralPath $chromeAppDir -Recurse -Force -ErrorAction SilentlyContinue)
        Write-Log 'INFO' "Application récursive des ACL sur $($items.Count) éléments..."
        foreach ($item in $items) {
            try {
                $itemAcl = Get-Acl -LiteralPath $item.FullName
                $itemAcl.SetOwner($adminAccount)
                Set-Acl -LiteralPath $item.FullName -AclObject $itemAcl -ErrorAction Stop
                $itemAcl = Get-Acl -LiteralPath $item.FullName
                $itemAcl.SetAccessRuleProtection($true, $false)
                foreach ($rule in @($itemAcl.Access)) { $itemAcl.RemoveAccessRule($rule) | Out-Null }
                if ($item.PSIsContainer) {
                    $itemAcl.AddAccessRule($adminRule)
                    $itemAcl.AddAccessRule($systemRule)
                }
                else {
                    $fileAdminRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                        $adminGroup, 'FullControl', 'None', 'None', 'Allow')
                    $fileSystemRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
                        'SYSTEM', 'FullControl', 'None', 'None', 'Allow')
                    $itemAcl.AddAccessRule($fileAdminRule)
                    $itemAcl.AddAccessRule($fileSystemRule)
                }
                Set-Acl -LiteralPath $item.FullName -AclObject $itemAcl -ErrorAction Stop
            }
            catch {
                $errors++
                Write-Log 'WARNING' "ACL non appliquée sur $($item.FullName) : $($_.Exception.Message)"
            }
        }
        if ($errors -eq 0) {
            Write-Log 'SUCCESS' "ACL appliquée récursivement sur $chromeAppDir : seuls $adminGroup + SYSTEM ont accès"
        }
        else {
            Write-Log 'WARNING' "ACL appliquée avec $errors erreurs sur $chromeAppDir"
        }
        # Vérification
        $testAcl = Get-Acl -LiteralPath $chromeAppDir
        $nonAdmin = $testAcl.Access | Where-Object { $_.IdentityReference -notmatch "SYSTEM|Syst.me|AUTORITE|$([regex]::Escape($adminGroup))|BUILTIN" }
        if ($nonAdmin) {
            Write-Log 'WARNING' "ACL vérification : des permissions non-admin subsistent sur $chromeAppDir"
            $nonAdmin | ForEach-Object { Write-Log 'WARNING' "  -> $($_.IdentityReference) : $($_.FileSystemRights)" }
        }
        else {
            Write-Log 'SUCCESS' "ACL vérification OK : seuls $adminGroup + SYSTEM ont accès à $chromeAppDir"
        }
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
$chromeExe = Get-ChromeExePath
if ($chromeExe) {
    Write-Log 'INFO' "Chrome trouvé : $chromeExe"
    Set-ChromeRestrictionAcl -ChromeExePath $chromeExe
} else {
    Write-Log 'INFO' "Chrome n''est pas installé sur ce poste."
}
Invoke-ChromeTaskbarAutoRepair
Write-Log 'SUCCESS' 'Traitement terminé.'
