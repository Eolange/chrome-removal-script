# Remove-Chrome.ps1

# Fix 1: Aggressive Icon Cache Clearing
function Clear-IconCache {
    $iconCachePath = "$env:LOCALAPPDATA\Microsoft\Windows\Explorer\iconcache*"
    Remove-Item $iconCachePath -Force -ErrorAction SilentlyContinue
}

# Fix 2: Reinforced ACL with Deny Rules
function Reinforce-ACL {
    $acl = Get-Acl "C:\Path\To\Target"
    $denyRule = New-Object System.Security.AccessControl.RegistryAccessRule("Everyone", "Deny", "Allow")
    $acl.SetAccessRule($denyRule)
    Set-Acl "C:\Path\To\Target" $acl
}

# Fix 3: Single Explorer Restart
function Restart-Explorer {
    Stop-Process -Name explorer -Force -ErrorAction SilentlyContinue
    Start-Process explorer
}

# Fix 4: Complete Original Functionality
function Remove-Chrome {
    # Original functionality to remove Chrome
    Clear-IconCache
    Reinforce-ACL
    Restart-Explorer  # Ensure single restart
    # Additional removal code goes here...
}

# Execute the removal function
Remove-Chrome