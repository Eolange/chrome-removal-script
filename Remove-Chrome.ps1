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

# Prise de possession + modification ACL en DEUX étapes (sinon Access Denied)
function Set-AclSafe {
    param(
        [string]$Path,
        [System.Security.Principal.NTAccount]$Owner,
        [System.Security.AccessControl.FileSystemAccessRule[]]$Rules
    )
    # Étape 1 : prendre possession
    $acl = Get-Acl -LiteralPath $Path
    $acl.SetOwner($Owner)
    Set-Acl -LiteralPath $Path -AclObject $acl -ErrorAction Stop

    # Étape 2 : relire l'ACL (en tant que propriétaire) puis modifier les règles
    $acl = Get-Acl -LiteralPath $Path
    $acl.SetAccessRuleProtection($true, $false)   # couper l'héritage, ne pas copier
    foreach ($r in @($acl.Access)) { $acl.RemoveAccessRule($r) | Out-Null }
    foreach ($r in $Rules) { $acl.AddAccessRule($r) }
    Set-Acl -LiteralPath $Path -AclObject $acl -ErrorAction Stop
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

#region 5. Bloquer chrome.exe par ACL + SRP
$chromePaths = @(
    "${env:ProgramFiles(x86)}\Google\Chrome\Application\chrome.exe",
    "${env:ProgramFiles}\Google\Chrome\Application\chrome.exe"
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) }

if (-not $chromePaths) {
    Write-Log INFO "Chrome n'est pas installé sur ce poste."
} else {
    # Activer les privilèges nécessaires pour prendre possession
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

    # Règles ACL pour les DOSSIERS (admins + SYSTEM full, Users lecture+listage)
    $dirRules = @(
        (New-Object System.Security.AccessControl.FileSystemAccessRule($AdminGroup, 'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')),
        (New-Object System.Security.AccessControl.FileSystemAccessRule('SYSTEM',     'FullControl', 'ContainerInherit,ObjectInherit', 'None', 'Allow')),
        (New-Object System.Security.AccessControl.FileSystemAccessRule($UsersGroup,  'ReadAndExecute,ListDirectory', 'ContainerInherit,ObjectInherit', 'None', 'Allow'))
    )

    # Règles ACL pour les EXE (admins + SYSTEM full, Users lecture SEULE = pas d'exécution)
    $exeRules = @(
        (New-Object System.Security.AccessControl.FileSystemAccessRule($AdminGroup, 'FullControl', 'None', 'None', 'Allow')),
        (New-Object System.Security.AccessControl.FileSystemAccessRule('SYSTEM',     'FullControl', 'None', 'None', 'Allow')),
        (New-Object System.Security.AccessControl.FileSystemAccessRule($UsersGroup,  'Read',        'None', 'None', 'Allow'))
    )

    foreach ($chromeExe in $chromePaths) {
        $chromeAppDir = Split-Path (Split-Path $chromeExe -Parent) -Parent
        Write-Log INFO "Blocage de : $chromeExe"

        try {
            # ACL dossier
            Set-AclSafe -Path $chromeAppDir -Owner $adminAccount -Rules $dirRules
            Write-Log SUCCESS "ACL dossier OK : $chromeAppDir"

            # ACL sur chaque .exe (bloquer exécution pour Users)
            foreach ($exe in (Get-ChildItem (Split-Path $chromeExe -Parent) -Filter '*.exe' -ErrorAction SilentlyContinue)) {
                Set-AclSafe -Path $exe.FullName -Owner $adminAccount -Rules $exeRules
                Write-Log SUCCESS "ACL exe bloqué : $($exe.Name)"
            }
        }
        catch {
            Write-Log ERROR "Blocage ACL échoué : $($_.Exception.Message)"
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
