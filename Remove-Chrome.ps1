#requires -RunAsAdministrator
<#
.SYNOPSIS
    Bloque Google Chrome pour les utilisateurs non-administrateurs.
.DESCRIPTION
    - Arrête les processus Chrome
    - Supprime les raccourcis (bureau, taskbar, menu démarrer)
    - Désactive les tâches planifiées et services Google Update
    - Bloque chrome.exe par ACL (lecture seule pour Users) + SRP
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

#region 2. Supprimer les raccourcis
$shortcutPaths = @(
    "$env:PUBLIC\Desktop\Google Chrome.lnk",
    "$env:ProgramData\Microsoft\Windows\Start Menu\Programs\Google Chrome.lnk"
)
# Ajouter les raccourcis de chaque profil utilisateur
Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $shortcutPaths += Join-Path $_.FullName 'Desktop\Google Chrome.lnk'
}

foreach ($path in $shortcutPaths) {
    if (Test-Path -LiteralPath $path) {
        Remove-Item -LiteralPath $path -Force -ErrorAction SilentlyContinue
        Write-Log SUCCESS "Raccourci supprimé : $path"
    }
}

# Nettoyer la taskbar et Quick Launch (tous les profils)
$wsh = New-Object -ComObject WScript.Shell -ErrorAction SilentlyContinue
Get-ChildItem 'C:\Users' -Directory -ErrorAction SilentlyContinue | ForEach-Object {
    $qlRoot = Join-Path $_.FullName 'AppData\Roaming\Microsoft\Internet Explorer\Quick Launch'
    if (-not (Test-Path -LiteralPath $qlRoot)) { return }

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
#endregion

#region 3. Désactiver tâches planifiées et services Google
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

# Supprimer Google Update
foreach ($dir in "${env:ProgramFiles(x86)}\Google\Update", "${env:ProgramFiles}\Google\Update") {
    if (Test-Path -LiteralPath $dir) {
        Remove-Item -LiteralPath $dir -Recurse -Force -ErrorAction SilentlyContinue
        Write-Log SUCCESS "Google Update supprimé : $dir"
    }
}
#endregion

#region 4. Bloquer chrome.exe par ACL + SRP
$chromePaths = @(
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe"
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if (-not $chromePaths) {
    Write-Log INFO "Chrome n'est pas installé sur ce poste."
} else {
    # Activer les privilèges nécessaires
    $tokenPrivCode = @'
using System;
using System.Runtime.InteropServices;
public class TokenPriv {
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool AdjustTokenPrivileges(IntPtr h, bool d, ref TP n, int l, IntPtr p, IntPtr r);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool OpenProcessToken(IntPtr h, uint a, out IntPtr t);
    [DllImport("advapi32.dll", SetLastError=true)]
    static extern bool LookupPrivilegeValue(string s, string n, out long l);
    struct TP { public int Count; public long Luid; public int Attr; }
    public static void Enable(string priv) {
        IntPtr token; long luid;
        OpenProcessToken((IntPtr)(-1), 0x28, out token);
        LookupPrivilegeValue(null, priv, out luid);
        TP tp = new TP { Count = 1, Luid = luid, Attr = 2 };
        AdjustTokenPrivileges(token, false, ref tp, 0, IntPtr.Zero, IntPtr.Zero);
    }
}
'@
    try { Add-Type $tokenPrivCode -ErrorAction SilentlyContinue } catch {}
    try { [TokenPriv]::Enable('SeTakeOwnershipPrivilege'); [TokenPriv]::Enable('SeRestorePrivilege') } catch {}

    $adminAccount = New-Object System.Security.Principal.NTAccount($AdminGroup)

    foreach ($chromeExe in $chromePaths) {
        $chromeAppDir = Split-Path (Split-Path $chromeExe -Parent) -Parent
        Write-Log INFO "Blocage de : $chromeExe"

        try {
            # ACL dossier : lisible par tous, modifiable par admins
            $dirAcl = Get-Acl -LiteralPath $chromeAppDir
            $dirAcl.SetOwner($adminAccount)
            $dirAcl.SetAccessRuleProtection($true, $false)
            $dirAcl.Access | ForEach-Object { $dirAcl.RemoveAccessRule($_) | Out-Null }
            $dirAcl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($AdminGroup, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')))
            $dirAcl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule('SYSTEM', 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')))
            $dirAcl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($UsersGroup, 'ReadAndExecute,ListDirectory', 'ContainerInherit,ObjectInherit', 'None', 'Allow')))
            Set-Acl -LiteralPath $chromeAppDir -AclObject $dirAcl
            Write-Log SUCCESS "ACL dossier : $chromeAppDir"

            # ACL exe : lecture seule pour Users (pas d'exécution)
            foreach ($exe in (Get-ChildItem (Split-Path $chromeExe -Parent) -Filter '*.exe' -ErrorAction SilentlyContinue)) {
                $exeAcl = Get-Acl -LiteralPath $exe.FullName
                $exeAcl.SetOwner($adminAccount)
                $exeAcl.SetAccessRuleProtection($true, $false)
                $exeAcl.Access | ForEach-Object { $exeAcl.RemoveAccessRule($_) | Out-Null }
                $exeAcl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($AdminGroup, 'FullControl', 'None', 'None', 'Allow')))
                $exeAcl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule('SYSTEM', 'FullControl', 'None', 'None', 'Allow')))
                $exeAcl.AddAccessRule((New-Object System.Security.AccessControl.FileSystemAccessRule($UsersGroup, 'Read', 'None', 'None', 'Allow')))
                Set-Acl -LiteralPath $exe.FullName -AclObject $exeAcl
                Write-Log SUCCESS "ACL exe bloqué : $($exe.Name)"
            }
        }
        catch {
            Write-Log ERROR "Blocage ACL échoué : $($_.Exception.Message)"
        }
    }

    # SRP (Software Restriction Policy)
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
