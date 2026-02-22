function Test-PnPModuleAvailable {
    
    # Method 1: Check installed modules (prefer 3.x)
    $moduleInstalled = Get-Module -ListAvailable -Name PnP.PowerShell -ErrorAction SilentlyContinue |
        Sort-Object Version -Descending | Select-Object -First 1
    
    if ($moduleInstalled) {
        $version = $moduleInstalled.Version
        Write-ActivityLog "PnP module found via Get-Module: Version $version"
        
        # Check if it's modern version
        if ($version -ge [version]"3.0.0") {
            Write-ActivityLog "Modern PnP PowerShell 3.x detected"
            return $true
        } elseif ($version -ge [version]"2.0.0") {
            Write-ActivityLog "Legacy PnP PowerShell 2.x detected - recommend upgrading to 3.x"
            return $true
        } else {
            Write-ActivityLog "Very old PnP PowerShell detected - upgrade required"
            return $false
        }
    }
    
    # Method 2: Try to import
    try {
        Import-Module PnP.PowerShell -ErrorAction Stop
        $importedModule = Get-Module PnP.PowerShell
        if ($importedModule) {
            Write-ActivityLog "PnP module successfully imported: Version $($importedModule.Version)"
            return $true
        }
    }
    catch {
        Write-ActivityLog "PnP module import failed: $($_.Exception.Message)"
    }
    
    # Method 3: Check if key commands are available (modern commands)
    try {
        $modernCommands = @("Connect-PnPOnline", "Get-PnPWeb", "Get-PnPProperty")
        $availableCommands = 0
        
        foreach ($cmd in $modernCommands) {
            if (Get-Command $cmd -ErrorAction SilentlyContinue) {
                $availableCommands++
            }
        }
        
        if ($availableCommands -eq $modernCommands.Count) {
            Write-ActivityLog "All modern PnP commands are available"
            return $true
        } else {
            Write-ActivityLog "Missing $($modernCommands.Count - $availableCommands) key PnP commands"
        }
    }
    catch {
        Write-ActivityLog "PnP commands check failed: $($_.Exception.Message)"
    }
    
    return $false
}

function Test-UserCapabilities {
    <#
    .SYNOPSIS
    Tests what SharePoint and Graph capabilities the connected user has access to

    .DESCRIPTION
    Returns a hashtable of capability flags indicating what operations the user can perform:
    - CanEnumerateSites: Can call Get-PnPTenantSite (requires SharePoint Admin)
    - CanReadUsers: Can call Graph /v1.0/users (requires User.Read.All)
    - CanAccessStorageData: Can get detailed storage information (requires SharePoint Admin)
    - CanReadExternalUsers: Can query external user data (requires User.Read.All)
    #>
    [CmdletBinding()]
    param()

    $capabilities = @{
        CanEnumerateSites = $false
        CanReadUsers = $false
        CanAccessStorageData = $false
        CanReadExternalUsers = $false
        CheckedAt = (Get-Date).ToString("o")
    }

    # Test 1: Site Enumeration (requires SharePoint Admin)
    try {
        Write-ActivityLog "Testing site enumeration capability..." -Level "Information"
        # Test permission by attempting the call - success means we have permission
        # Don't check if data is returned, check if the call succeeds
        $null = Get-PnPTenantSite -ErrorAction Stop
        $capabilities.CanEnumerateSites = $true
        Write-ActivityLog "Site enumeration: ENABLED" -Level "Information"
    }
    catch {
        Write-ActivityLog "Site enumeration: DISABLED - $($_.Exception.Message)" -Level "Information"
        $capabilities.CanEnumerateSites = $false
    }

    # Small delay to avoid throttling
    Start-Sleep -Milliseconds 500

    # Test 2: User Read via Graph (requires User.Read.All)
    try {
        Write-ActivityLog "Testing Graph user read capability..." -Level "Information"
        # Test permission by attempting the call - success means we have permission
        $null = Invoke-PnPGraphMethod -Url "v1.0/users?`$top=1" -Method Get -ErrorAction Stop
        $capabilities.CanReadUsers = $true
        Write-ActivityLog "User iteration: ENABLED" -Level "Information"
    }
    catch {
        Write-ActivityLog "User iteration: DISABLED - $($_.Exception.Message)" -Level "Information"
        $capabilities.CanReadUsers = $false
    }

    # Small delay to avoid throttling
    Start-Sleep -Milliseconds 500

    # Test 3: Storage Data (requires SharePoint Admin or Sites.FullControl.All)
    try {
        Write-ActivityLog "Testing storage data access..." -Level "Information"
        # Test permission by attempting the call with -Detailed flag
        $null = Get-PnPTenantSite -Detailed -ErrorAction Stop
        $capabilities.CanAccessStorageData = $true
        Write-ActivityLog "Storage data: ENABLED" -Level "Information"
    }
    catch {
        Write-ActivityLog "Storage data: LIMITED - $($_.Exception.Message)" -Level "Information"
        $capabilities.CanAccessStorageData = $false
    }

    # Small delay to avoid throttling
    Start-Sleep -Milliseconds 500

    # Test 4: External User Data (requires User.Read.All)
    try {
        Write-ActivityLog "Testing external user data access..." -Level "Information"
        # Test permission by attempting the filtered query
        $null = Invoke-PnPGraphMethod -Url "v1.0/users?`$filter=userType eq 'Guest'&`$top=1" -Method Get -ErrorAction Stop
        $capabilities.CanReadExternalUsers = $true
        Write-ActivityLog "External user data: ENABLED" -Level "Information"
    }
    catch {
        Write-ActivityLog "External user data: DISABLED - $($_.Exception.Message)" -Level "Information"
        $capabilities.CanReadExternalUsers = $false
    }

    return $capabilities
}
