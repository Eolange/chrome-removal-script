function Reset-UserIconCacheAggressive { 
    $iconCachePath = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer" 
    $iconCacheFile = "iconcache*" 
    Remove-Item "$iconCachePath\$iconCacheFile" -ErrorAction SilentlyContinue 
    Stop-Process -Name "explorer" -Force -ErrorAction SilentlyContinue 
    Start-Sleep -Seconds 2 
    Start-Process "explorer.exe" 
}

function Set-ChromeRestrictionAcl { 
    $chromePath = "C:\Program Files\Google\Chrome\Application" 
    $acl = Get-Acl $chromePath 
    $denyRules = @(
        "RX", "W", "D", "DC"
    ) 
    foreach ($rule in $denyRules) { 
        $accessRule = New-Object System.Security.AccessControl.FileSystemAccessRule("Everyone", $rule, "Deny") 
        $acl.SetAccessRule($accessRule) 
    } 
    Set-Acl $chromePath $acl 
}

function Invoke-ChromeTaskbarAutoRepair { 
    $taskbarPath = "$env:APPDATA\Microsoft\Internet Explorer\Quick Launch\Taskbar" 
    $chromeShortcut = "Google Chrome.lnk" 
    if (Test-Path "$taskbarPath\$chromeShortcut") { 
        Remove-Item "$taskbarPath\$chromeShortcut" -ErrorAction SilentlyContinue 
        Start-Sleep -Seconds 1 
    } 
    $chromeExePath = "C:\Program Files\Google\Chrome\Application\chrome.exe" 
    Start-Process "$chromeExePath" 
}

# Reset user icon cache aggressively
Reset-UserIconCacheAggressive 

# Set Chrome restriction ACL
Set-ChromeRestrictionAcl 

# Invoke Chrome taskbar auto-repair without explorer restart
Invoke-ChromeTaskbarAutoRepair 

# Restart Explorer once at the end
Stop-Process -Name "explorer" -Force 
Start-Sleep -Seconds 1 
Start-Process "explorer.exe" 
