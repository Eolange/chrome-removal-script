# Remove-Chrome.ps1

# A PowerShell script to remove Google Chrome and clean up caches.

function Reset-UserIconCacheAggressive {
    # Reset user icon cache aggressively
    Stop-Process -Name explorer -Force
    Remove-Item -Path "$env:APPDATA\\Microsoft\\Windows\\Explorer\\iconcache*" -ErrorAction SilentlyContinue
    Start-Sleep -Seconds 2
    Start-Process explorer
}

function Set-ChromeRestrictionAcl {
    param (
        [string]$path
    )

    # Explicitly DENY rules to restrict access to Chrome directories
    $acl = Get-Acl $path
    $denyRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", "FullControl", "Deny")
    $acl.AddAccessRule($denyRule)
    Set-Acl $path $acl
}

function Remove-Chrome {
    # Remove Google Chrome application
    Get-Process chrome -ErrorAction SilentlyContinue | Stop-Process -Force
    Start-Sleep -Seconds 2

    # Uninstall Chrome using WMIC
    wmic product where "name='Google Chrome'" call uninstall /nointeractive
    Start-Sleep -Seconds 5

    # Clean up application data
    Remove-Item -Path "$env:LOCALAPPDATA\\Google\\Chrome" -Recurse -Force -ErrorAction SilentlyContinue
    Remove-Item -Path "$env:APPDATA\\Google\\Chrome" -Recurse -Force -ErrorAction SilentlyContinue

    # Aggressively clean cache
    Get-ChildItem -Path "$env:TEMP" -Recurse | Remove-Item -Force -ErrorAction SilentlyContinue
    Reset-UserIconCacheAggressive
}

# Main script execution
Remove-Chrome
