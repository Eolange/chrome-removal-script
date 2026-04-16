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
    $profiles = Get-ActiveUserProfiles
    $wshShell = $null
    try { $wshShell = New-Object -ComObject WScript.Shell } catch {}
    $totalRemoved = 0
    # 1. Supprimer tous les raccourcis Chrome résiduels
    foreach ($profile in $profiles) {
        $quickLaunchRoot = Join-Path $profile.LocalPath 'AppData\Roaming\Microsoft\Internet Explorer\Quick Launch'
        if (-not (Test-Path -LiteralPath $quickLaunchRoot)) { continue }
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
    # 2. Tenter un désépinglage immédiat via Register-ScheduledTask (cmdlet PS, pas schtasks.exe)
    $immediateOk = $false
    try {
        $computerSystem = Get-CimInstance Win32_ComputerSystem -ErrorAction SilentlyContinue
        $interactiveUser = $computerSystem.UserName
        if ($interactiveUser) {
            $userNameOnly = $interactiveUser.Split('\')[-1]
            $userProfile = $profiles | Where-Object { $_.LocalPath -match [regex]::Escape($userNameOnly) } | Select-Object -First 1
            if ($userProfile) {
                $vbsPath = Join-Path $userProfile.LocalPath 'AppData\Local\Temp\UnpinChrome.vbs'
                $vbsContent = @"
Set sh = CreateObject(""Shell.Application"")
Set wsh = CreateObject(""WScript.Shell"")
tb = wsh.ExpandEnvironmentStrings(""%APPDATA%"") & ""\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar""
Set fso = CreateObject(""Scripting.FileSystemObject"")
If fso.FolderExists(tb) Then
    Set folder = sh.Namespace(tb)
    If Not folder Is Nothing Then
        For Each item In folder.Items
            nm = LCase(item.Name)
            If InStr(nm, ""chrome"") > 0 Then
                For Each verb In item.Verbs
                    If InStr(verb.Name, ""pingler"") > 0 Or InStr(verb.Name, ""Unpin"") > 0 Then
                        verb.DoIt
                    End If
                Next
            End If
        Next
    End If
End If
fso.DeleteFile WScript.ScriptFullName, True
"@
                Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII -Force
                $taskName = 'UnpinChromeImmediate'
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
                $action = New-ScheduledTaskAction -Execute 'wscript.exe' -Argument "`"$vbsPath`""
                $principal = New-ScheduledTaskPrincipal -UserId $interactiveUser -LogonType Interactive -RunLevel Limited
                $settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries
                Register-ScheduledTask -TaskName $taskName -Action $action -Principal $principal -Settings $settings -Force -ErrorAction Stop | Out-Null
                Start-ScheduledTask -TaskName $taskName -ErrorAction Stop
                Write-Log 'SUCCESS' "Désépinglage immédiat lancé dans la session de $interactiveUser"
                $immediateOk = $true
                Start-Sleep -Seconds 3
                Unregister-ScheduledTask -TaskName $taskName -Confirm:$false -ErrorAction SilentlyContinue
            }
        }
    }
    catch {
        Write-Log 'WARNING' "Désépinglage immédiat impossible : $($_.Exception.Message)"
    }
    # 3. Fallback : enregistrer RunOnce pour le prochain logon
    if (-not $immediateOk) {
        foreach ($profile in $profiles) {
            if (-not $profile.Loaded) { continue }
            $sid = $profile.SID
            $runOncePath = "Registry::HKU\$sid\Software\Microsoft\Windows\CurrentVersion\RunOnce"
            try {
                if (-not (Test-Path $runOncePath)) {
                    New-Item -Path $runOncePath -Force -ErrorAction Stop | Out-Null
                }
                $vbsPath = Join-Path $profile.LocalPath 'AppData\Local\Temp\UnpinChrome.vbs'
                $vbsContent = @"
Set sh = CreateObject(""Shell.Application"")
Set wsh = CreateObject(""WScript.Shell"")
tb = wsh.ExpandEnvironmentStrings(""%APPDATA%"") & ""\Microsoft\Internet Explorer\Quick Launch\User Pinned\TaskBar""
Set fso = CreateObject(""Scripting.FileSystemObject"")
If fso.FolderExists(tb) Then
    Set folder = sh.Namespace(tb)
    If Not folder Is Nothing Then
        For Each item In folder.Items
            nm = LCase(item.Name)
            If InStr(nm, ""chrome"") > 0 Then
                For Each verb In item.Verbs
                    If InStr(verb.Name, ""pingler"") > 0 Or InStr(verb.Name, ""Unpin"") > 0 Then
                        verb.DoIt
                    End If
                Next
            End If
        Next
    End If
End If
fso.DeleteFile WScript.ScriptFullName, True
"@
                Set-Content -Path $vbsPath -Value $vbsContent -Encoding ASCII -Force
                Set-ItemProperty -Path $runOncePath -Name 'UnpinChrome' -Value "wscript.exe `"$vbsPath`"" -Type String -ErrorAction Stop
                Write-Log 'SUCCESS' "RunOnce enregistré : l''icône Chrome sera supprimée au prochain logon (SID: $sid)"
            }
            catch {
                Write-Log 'WARNING' "Impossible d''enregistrer RunOnce pour SID $sid : $($_.Exception.Message)"
            }
        }
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
    $usersGroup = if (Test-IdentityResolvable -Identity 'Utilisateurs') { 'Utilisateurs' } else { 'Users' }
    Write-Log 'INFO' "Blocage ACL Chrome : chrome.exe uniquement (dossier reste lisible)"
    Write-Log 'INFO' "Groupe admin détecté : $adminGroup / Groupe users : $usersGroup"
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
        # --- DOSSIER : lisible par tout le monde, modifiable par admins ---
        $acl = Get-Acl -LiteralPath $chromeAppDir
        $acl.SetOwner($adminAccount)
        Set-Acl -LiteralPath $chromeAppDir -AclObject $acl -ErrorAction Stop
        $acl = Get-Acl -LiteralPath $chromeAppDir
        $acl.SetAccessRuleProtection($true, $false)
        $acl.Access | ForEach-Object { $acl.RemoveAccessRule($_) | Out-Null }
        $adminRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $adminGroup, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $systemRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            'SYSTEM', 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $usersRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $usersGroup, 'ReadAndExecute,ListDirectory', 'ContainerInherit,ObjectInherit', 'None', 'Allow')
        $acl.AddAccessRule($adminRule)
        $acl.AddAccessRule($systemRule)
        $acl.AddAccessRule($usersRule)
        Set-Acl -LiteralPath $chromeAppDir -AclObject $acl -ErrorAction Stop
        Write-Log 'SUCCESS' "ACL dossier appliquée sur $chromeAppDir ($adminGroup + SYSTEM + $usersGroup en lecture)"
        # --- CHROME.EXE : bloquer l'exécution pour les non-admins ---
        $exeAcl = Get-Acl -LiteralPath $ChromeExePath
        $exeAcl.SetOwner($adminAccount)
        Set-Acl -LiteralPath $ChromeExePath -AclObject $exeAcl -ErrorAction Stop
        $exeAcl = Get-Acl -LiteralPath $ChromeExePath
        $exeAcl.SetAccessRuleProtection($true, $false)
        foreach ($rule in @($exeAcl.Access)) { $exeAcl.RemoveAccessRule($rule) | Out-Null }
        $exeAdminRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $adminGroup, 'FullControl', 'None', 'None', 'Allow')
        $exeSystemRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            'SYSTEM', 'FullControl', 'None', 'None', 'Allow')
        $exeUsersReadRule = New-Object System.Security.AccessControl.FileSystemAccessRule(
            $usersGroup, 'Read', 'None', 'None', 'Allow')
        $exeAcl.AddAccessRule($exeAdminRule)
        $exeAcl.AddAccessRule($exeSystemRule)
        $exeAcl.AddAccessRule($exeUsersReadRule)
        Set-Acl -LiteralPath $ChromeExePath -AclObject $exeAcl -ErrorAction Stop
        Write-Log 'SUCCESS' "ACL chrome.exe : $adminGroup + SYSTEM FullControl, $usersGroup Read seul (pas d''exécution)"
        # Bloquer aussi new_chrome.exe et chrome_proxy.exe s'ils existent
        $otherExes = @('new_chrome.exe', 'chrome_proxy.exe')
        foreach ($exeName in $otherExes) {
            $otherExePath = Join-Path $chromeDir $exeName
            if (Test-Path -LiteralPath $otherExePath) {
                try {
                    $otherAcl = Get-Acl -LiteralPath $otherExePath
                    $otherAcl.SetOwner($adminAccount)
                    Set-Acl -LiteralPath $otherExePath -AclObject $otherAcl -ErrorAction Stop
                    $otherAcl = Get-Acl -LiteralPath $otherExePath
                    $otherAcl.SetAccessRuleProtection($true, $false)
                    foreach ($rule in @($otherAcl.Access)) { $otherAcl.RemoveAccessRule($rule) | Out-Null }
                    $otherAcl.AddAccessRule($exeAdminRule)
                    $otherAcl.AddAccessRule($exeSystemRule)
                    $otherAcl.AddAccessRule($exeUsersReadRule)
                    Set-Acl -LiteralPath $otherExePath -AclObject $otherAcl -ErrorAction Stop
                    Write-Log 'SUCCESS' "ACL bloquée sur $exeName"
                }
                catch {
                    Write-Log 'WARNING' "ACL non appliquée sur $otherExePath : $($_.Exception.Message)"
                }
            }
        }
        # Vérification sur chrome.exe
        $testAcl = Get-Acl -LiteralPath $ChromeExePath
        $nonAdmin = $testAcl.Access | Where-Object { $_.IdentityReference -notmatch "SYSTEM|Syst.me|AUTORITE|$([regex]::Escape($adminGroup))|$([regex]::Escape($usersGroup))|BUILTIN" }
        if ($nonAdmin) {
            Write-Log 'WARNING' "ACL vérification chrome.exe : des permissions inattendues"
            $nonAdmin | ForEach-Object { Write-Log 'WARNING' "  -> $($_.IdentityReference) : $($_.FileSystemRights)" }
        }
        else {
            Write-Log 'SUCCESS' "ACL vérification OK : chrome.exe ($adminGroup + SYSTEM FullControl, $usersGroup Read)"
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
