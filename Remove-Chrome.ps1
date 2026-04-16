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
    $profiles = Get-ActiveUserProfiles
    Write-Log 'INFO' 'Arrêt de l''explorateur et des processus Shell pour nettoyage.'
    foreach ($procName in @('explorer', 'ShellExperienceHost', 'SearchHost', 'SearchApp', 'StartMenuExperienceHost')) {
        $proc = Get-Process $procName -ErrorAction SilentlyContinue
        if ($proc) {
            $proc | Stop-Process -Force -ErrorAction SilentlyContinue
            Write-Log 'INFO' "Processus arrêté : $procName"
        }
    }
    Start-Sleep -Seconds 3
    foreach ($profile in $profiles) {
        if ($profile.Loaded -or (Test-Path -LiteralPath $profile.LocalPath)) {
            Reset-UserIconCache -UserProfilePath $profile.LocalPath
        }
    }
    Start-Process explorer.exe -ErrorAction SilentlyContinue
    Write-Log 'INFO' 'Explorateur redémarré après nettoyage complet.'
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
        $items | ForEach-Object {
            try {
                $itemAcl = Get-Acl -LiteralPath $_.FullName
                $itemAcl.SetOwner($adminAccount)
                Set-Acl -LiteralPath $_.FullName -AclObject $itemAcl -ErrorAction Stop
                $itemAcl = Get-Acl -LiteralPath $_.FullName
                $itemAcl.SetAccessRuleProtection($true, $false)
                $itemAcl.Access | ForEach-Object { $itemAcl.RemoveAccessRule($_) | Out-Null }
                $itemAcl.AddAccessRule($adminRule)
                $itemAcl.AddAccessRule($systemRule)
                Set-Acl -LiteralPath $_.FullName -AclObject $itemAcl -ErrorAction Stop
            }
            catch {
                $errors++
                Write-Log 'WARNING' "ACL non appliquée sur $($_.FullName) : $($_.Exception.Message)"
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
        $nonAdmin = $testAcl.Access | Where-Object { $_.IdentityReference -notmatch "SYSTEM|$([regex]::Escape($adminGroup))|BUILTIN" }
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
