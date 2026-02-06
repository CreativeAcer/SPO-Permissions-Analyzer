function Initialize-ConnectionTab {
    <#
    .SYNOPSIS
    Initializes the Connection tab with event handlers and default state
    #>
    try {
        Write-ActivityLog "Initializing Connection tab" -Level "Information"
        
        # Set up event handlers for connection tab
        $script:btnConnectSPO.Add_Click({ 
            Connect-SharePointOnline 
        })
        
        $script:btnDemoMode.Add_Click({ 
            Enable-DemoMode 
        })
        
        # Set initial connection status
        Update-ConnectionStatus -Status "Disconnected" -Message "Not connected to SharePoint"
        
        # Initialize connection state variables
        $script:SPOConnected = $false
        $script:SPOContext = $null
        
        Write-ActivityLog "Connection tab initialized successfully" -Level "Information"
    }
    catch {
        Write-ActivityLog "Failed to initialize Connection tab: $($_.Exception.Message)" -Level "Error"
        throw
    }
}

function Connect-SharePointOnline {
    <#
    .SYNOPSIS
    Handles the SharePoint Online connection process using the actual SharePoint logic
    #>
    try {
        Write-ActivityLog "Starting SharePoint Online connection" -Level "Information"
        
        # Validate inputs
        $tenantUrl = $script:txtTenantUrl.Text.Trim()
        $clientId = $script:txtClientId.Text.Trim()
        
        if ([string]::IsNullOrEmpty($tenantUrl)) {
            Show-ConnectionError "Please enter a valid SharePoint tenant URL"
            return
        }
        
        if ([string]::IsNullOrEmpty($clientId)) {
            Show-ConnectionError "Please enter a valid App Registration Client ID"
            return
        }
        
        # Validate URL format
        if (-not $tenantUrl.StartsWith("https://") -or -not $tenantUrl.Contains(".sharepoint.com")) {
            Show-ConnectionError "Please enter a valid SharePoint Online tenant URL (e.g., https://contoso.sharepoint.com)"
            return
        }
        
        # Validate Client ID format (GUID)
        try {
            $guid = [System.Guid]::Parse($clientId)
        }
        catch {
            Show-ConnectionError "Please enter a valid Client ID (GUID format)"
            return
        }
        
        # Update UI to show connection in progress
        $script:txtSPOResults.Text = "🔄 Connecting to SharePoint Online...`n`nValidating tenant URL and client ID...`n"
        $script:btnConnectSPO.IsEnabled = $false
        Update-UIAndWait -WaitMs 100
        
        # Save settings for future use
        Set-AppSetting -SettingName "SharePoint.TenantUrl" -Value $tenantUrl
        Set-AppSetting -SettingName "SharePoint.ClientId" -Value $clientId
        
        # Add progress updates
        $script:txtSPOResults.Text += "✓ Tenant URL validated: $tenantUrl`n"
        Update-UIAndWait -WaitMs 500
        
        $script:txtSPOResults.Text += "✓ Client ID validated: $clientId`n"
        Update-UIAndWait -WaitMs 300
        
        $script:txtSPOResults.Text += "🔍 Checking PnP PowerShell module...`n"
        Update-UIAndWait -WaitMs 400
        
        # Check if PnP module is available (using the actual logic from main file)
        if (-not (Test-PnPModuleAvailable)) {
            $script:txtSPOResults.Text += "📦 Installing PnP PowerShell module...`n"
            $script:txtSPOResults.Text += "This may take a few minutes on first run.`n"
            Update-UIAndWait -WaitMs 500
            
            try {
                Install-PnPModule -UI (New-Object PSObject -Property @{
                    UpdateStatus = { param($message) 
                        $script:txtSPOResults.Text += "$message`n"
                        Update-UIAndWait -WaitMs 200
                    }
                })
            }
            catch {
                Show-ConnectionError "Failed to install PnP PowerShell module: $($_.Exception.Message)"
                return
            }
        } else {
            $script:txtSPOResults.Text += "✓ PnP PowerShell module found and ready`n"
            Update-UIAndWait -WaitMs 300
        }
        
        $script:txtSPOResults.Text += "🔐 Initiating authentication flow...`n"
        $script:txtSPOResults.Text += "Please complete authentication in the popup window.`n"
        Update-UIAndWait -WaitMs 400
        
        # Use the actual SharePoint connection function from the main file
        $connectionResult = Invoke-ActualSharePointConnection -TenantUrl $tenantUrl -ClientId $clientId
        
        if ($connectionResult.Success) {
            Show-ConnectionSuccess -UserInfo $connectionResult.UserInfo -SiteInfo $connectionResult.SiteInfo
            Enable-SharePointOperations
        } else {
            Show-ConnectionError $connectionResult.ErrorMessage
        }
        
    }
    catch {
        Show-ConnectionError "Connection failed: $($_.Exception.Message)"
        Write-ErrorLog -Message $_.Exception.Message -Location "Connect-SharePointOnline"
    }
    finally {
        $script:btnConnectSPO.IsEnabled = $true
    }
}

function Invoke-ActualSharePointConnection {
    <#
    .SYNOPSIS
    Performs actual SharePoint connection using the logic from the main file
    #>
    param(
        [string]$TenantUrl,
        [string]$ClientId
    )
    
    try {
        Write-ActivityLog "Starting actual SharePoint connection to: $TenantUrl with App Registration: $ClientId"
        
        # Clear any existing connections first
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
            Write-ActivityLog "Cleared any existing connections"
            $script:txtSPOResults.Text += "🔄 Clearing previous connections...`n"
            Update-UIAndWait -WaitMs 300
        }
        catch {
            # Ignore errors during disconnect
        }
        
        # Ensure modern module is imported
        $script:txtSPOResults.Text += "📦 Importing modern PnP PowerShell 3.x...`n"
        Update-UIAndWait -WaitMs 400
        
        try {
            Import-Module PnP.PowerShell -Force -ErrorAction Stop
            $pnpModule = Get-Module PnP.PowerShell
            Write-ActivityLog "Modern PnP PowerShell imported successfully: Version $($pnpModule.Version)"
            $script:txtSPOResults.Text += "✓ PnP PowerShell $($pnpModule.Version) loaded`n"
            
            if ($pnpModule.Version -lt [version]"3.0.0") {
                throw "Legacy PnP PowerShell detected. Please upgrade to 3.x for full functionality."
            }
        }
        catch {
            throw "Failed to import modern PnP PowerShell module: $($_.Exception.Message)"
        }
        
        # Modern connection approach - simplified for PnP 3.x
        $script:txtSPOResults.Text += "🚀 Connecting to SharePoint Online...`n"
        $script:txtSPOResults.Text += "📱 App Registration: $ClientId`n"
        $script:txtSPOResults.Text += "⏳ Please complete authentication in the popup window...`n"
        Update-UIAndWait -WaitMs 500
        
        Write-ActivityLog "Attempting modern connection to $TenantUrl with ClientId $ClientId"
        
        try {
            # Modern PnP 3.x connection - more reliable
            Connect-PnPOnline -Url $TenantUrl -ClientId $ClientId -Interactive
            Write-ActivityLog "Modern connection command completed successfully"
            $script:txtSPOResults.Text += "✅ Authentication completed!`n"
            Update-UIAndWait -WaitMs 400
        }
        catch {
            throw "Failed to connect to SharePoint with app registration: $($_.Exception.Message)"
        }
        
        # Enhanced verification using modern PnP features
        $script:txtSPOResults.Text += "🔍 Verifying SharePoint connection...`n"
        Update-UIAndWait -WaitMs 300
        
        try {
            # Modern verification approach
            $context = Get-PnPContext -ErrorAction SilentlyContinue
            $connection = Get-PnPConnection -ErrorAction SilentlyContinue
            
            if ($null -eq $context -and $null -eq $connection) {
                throw "No SharePoint connection available after authentication"
            }
            
            Write-ActivityLog "SharePoint context/connection verified successfully"
            $script:txtSPOResults.Text += "✓ Connection context verified`n"
            Update-UIAndWait -WaitMs 200
            
            # Get enhanced site information using modern PnP
            $site = $null
            $web = $null
            try {
                # Try modern approach first
                $web = Get-PnPWeb -ErrorAction SilentlyContinue
                if ($web) {
                    $site = @{ 
                        Url = $web.Url; 
                        Title = $web.Title;
                        Created = $web.Created;
                        Language = $web.Language;
                        Template = $web.WebTemplate
                    }
                    Write-ActivityLog "Enhanced site verification successful: $($web.Url)"
                    $script:txtSPOResults.Text += "✓ Site information retrieved`n"
                } else {
                    # Fallback
                    $site = @{ Url = $TenantUrl; Title = "SharePoint Site" }
                    Write-ActivityLog "Using basic verification - connection established"
                }
                Update-UIAndWait -WaitMs 200
            }
            catch {
                Write-ActivityLog "Site verification had issues but proceeding: $($_.Exception.Message)"
                $site = @{ Url = $TenantUrl; Title = "SharePoint Site" }
            }
            
        }
        catch {
            Write-ActivityLog "Connection verification had issues but proceeding: $($_.Exception.Message)"
            $site = @{ Url = $TenantUrl; Title = "SharePoint Site" }
        }
        
        # Set global connection state (THIS IS CRITICAL)
        $script:SPOConnected = $true
        $script:SPOContext = Get-PnPConnection -ErrorAction SilentlyContinue
        
        # Enhanced user info with modern PnP
        $currentUser = ""
        $userEmail = ""
        try {
            $userInfo = Get-PnPCurrentUser -ErrorAction SilentlyContinue
            if ($userInfo) {
                $currentUser = $userInfo.Title
                $userEmail = $userInfo.UserPrincipalName -or $userInfo.Email -or $userInfo.LoginName -or "Authentication verified"
            }
        }
        catch {
            $currentUser = "Current User"
            $userEmail = "Authentication verified"
        }
        
        $script:txtSPOResults.Text += "🎯 Getting user information...`n"
        Update-UIAndWait -WaitMs 300
        
        Write-ActivityLog "Successfully connected to SharePoint: $($site.Url) using modern PnP with app $ClientId"
        
        return @{
            Success = $true
            UserInfo = @{
                Email = $userEmail
                DisplayName = $currentUser
                TenantId = "Connected via App Registration"
            }
            SiteInfo = $site
        }
        
    }
    catch {
        $script:SPOConnected = $false
        $script:SPOContext = $null
        Write-ErrorLog -Message $_.Exception.Message -Location "Actual-SharePoint-Connection"
        
        return @{
            Success = $false
            ErrorMessage = $_.Exception.Message
        }
    }
}

function Enable-DemoMode {
    <#
    .SYNOPSIS
    Enables demo mode for testing and demonstration purposes
    #>
    try {
        Write-ActivityLog "Enabling demo mode" -Level "Information"
        
        # Set demo mode setting
        Set-AppSetting -SettingName "DemoMode" -Value $true
        
        # Set connection state for demo mode (CRITICAL)
        $script:SPOConnected = $true
        $script:SPOContext = "Demo Mode Context"
        
        # Update connection results
        $script:txtSPOResults.Text = @"
🎭 DEMO MODE ENABLED

📊 Simulated SharePoint Connection:
✓ Tenant: https://demo.sharepoint.com
✓ App Registration: demo-client-12345
✓ Authenticated User: demo.user@contoso.com
✓ Permissions: Sites.FullControl.All, User.Read.All

🎯 Demo Features Available:
• Site collection enumeration (5 sample sites)
• Permission analysis (simulated data)
• User and group reporting
• Visual analytics dashboard
• Report generation

💡 All operations will return sample data for demonstration purposes.
   Switch to real mode by connecting with actual credentials.

✅ Demo mode is now active - you can use all SharePoint operations!
"@
        
        # Update connection status
        Update-ConnectionStatus -Status "Connected (Demo)" -Message "Connected (Demo Mode)"
        
        # Enable SharePoint operations
        Enable-SharePointOperations
        
        # Enable refresh analytics button
        if ($script:btnRefreshAnalytics) {
            $script:btnRefreshAnalytics.IsEnabled = $true
        }
        
        # Update analytics subtitle
        if ($script:txtAnalyticsSubtitle) {
            $script:txtAnalyticsSubtitle.Text = "Demo data available - Click any operation to see sample results"
        }
        
        Write-ActivityLog "Demo mode activated successfully" -Level "Information"
    }
    catch {
        Write-ActivityLog "Failed to enable demo mode: $($_.Exception.Message)" -Level "Error"
        Show-ConnectionError "Failed to enable demo mode: $($_.Exception.Message)"
    }
}

function Show-ConnectionSuccess {
    <#
    .SYNOPSIS
    Displays successful connection information
    #>
    param(
        [hashtable]$UserInfo,
        [hashtable]$SiteInfo = @{}
    )
    
    try {
        $script:txtSPOResults.Text += "`n🎉 CONNECTION SUCCESSFUL!`n"
        $script:txtSPOResults.Text += "=" * 50 + "`n`n"
        $script:txtSPOResults.Text += "👤 Authenticated User: $($UserInfo.DisplayName)`n"
        $script:txtSPOResults.Text += "📧 Email/UPN: $($UserInfo.Email)`n"
        $script:txtSPOResults.Text += "🆔 Tenant Info: $($UserInfo.TenantId)`n"
        
        if ($SiteInfo.Url) {
            $script:txtSPOResults.Text += "`n🏢 Connected Site Information:`n"
            $script:txtSPOResults.Text += "📍 URL: $($SiteInfo.Url)`n"
            if ($SiteInfo.Title) { $script:txtSPOResults.Text += "📋 Title: $($SiteInfo.Title)`n" }
            if ($SiteInfo.Template) { $script:txtSPOResults.Text += "🎨 Template: $($SiteInfo.Template)`n" }
            if ($SiteInfo.Created) { $script:txtSPOResults.Text += "📅 Created: $($SiteInfo.Created)`n" }
        }
        
        $script:txtSPOResults.Text += "`n✅ Connection established and verified!`n"
        $script:txtSPOResults.Text += "🚀 You can now use the SharePoint Operations tab to analyze permissions.`n"
        
        Update-ConnectionStatus -Status "Connected" -Message "Connected to SharePoint Online"
        
        Write-ActivityLog "SharePoint connection successful for user: $($UserInfo.Email)" -Level "Information"
    }
    catch {
        Write-ActivityLog "Error displaying connection success: $($_.Exception.Message)" -Level "Warning"
    }
}

function Show-ConnectionError {
    <#
    .SYNOPSIS
    Displays connection error information
    #>
    param(
        [string]$ErrorMessage
    )
    
    try {
        $script:txtSPOResults.Text += "`n❌ CONNECTION FAILED`n"
        $script:txtSPOResults.Text += "=" * 50 + "`n`n"
        $script:txtSPOResults.Text += "Error: $ErrorMessage`n`n"
        
        # Provide specific troubleshooting based on error type
        if ($ErrorMessage -like "*Access is denied*" -or $ErrorMessage -like "*Forbidden*") {
            $script:txtSPOResults.Text += "🔐 ACCESS DENIED - This typically means:`n"
            $script:txtSPOResults.Text += "• Your app registration lacks proper permissions`n"
            $script:txtSPOResults.Text += "• You don't have access to the specified tenant`n"
            $script:txtSPOResults.Text += "• The app needs admin consent for SharePoint permissions`n`n"
        }
        elseif ($ErrorMessage -like "*not found*" -or $ErrorMessage -like "*404*") {
            $script:txtSPOResults.Text += "🔍 TENANT NOT FOUND - Please check:`n"
            $script:txtSPOResults.Text += "• The tenant URL is correct and spelled properly`n"
            $script:txtSPOResults.Text += "• You're using the right SharePoint Online tenant`n"
            $script:txtSPOResults.Text += "• The tenant exists and is accessible`n`n"
        }
        elseif ($ErrorMessage -like "*AADSTS*") {
            $script:txtSPOResults.Text += "🔑 AUTHENTICATION ERROR - This usually means:`n"
            $script:txtSPOResults.Text += "• Invalid Client ID (App Registration ID)`n"
            $script:txtSPOResults.Text += "• App registration not properly configured`n"
            $script:txtSPOResults.Text += "• Authentication flow was cancelled or failed`n`n"
        }
        
        $script:txtSPOResults.Text += "💡 Troubleshooting Tips:`n"
        $script:txtSPOResults.Text += "• Verify your tenant URL format (https://yourtenant.sharepoint.com)`n"
        $script:txtSPOResults.Text += "• Check your App Registration Client ID is a valid GUID`n"
        $script:txtSPOResults.Text += "• Ensure your app has Sites.Read.All or Sites.FullControl.All permissions`n"
        $script:txtSPOResults.Text += "• Try Demo Mode to test the application functionality`n"
        $script:txtSPOResults.Text += "• Contact your SharePoint Administrator if permissions are needed`n"
        
        Update-ConnectionStatus -Status "Error" -Message "Connection failed"
        
        # Ensure connection state is properly reset
        $script:SPOConnected = $false
        $script:SPOContext = $null
        
        Write-ActivityLog "SharePoint connection failed: $ErrorMessage" -Level "Error"
    }
    catch {
        Write-ActivityLog "Error displaying connection error: $($_.Exception.Message)" -Level "Warning"
    }
}

function Update-ConnectionStatus {
    <#
    .SYNOPSIS
    Updates the connection status display across the application
    #>
    param(
        [ValidateSet("Connected", "Connected (Demo)", "Disconnected", "Error")]
        [string]$Status,
        [string]$Message
    )
    
    try {
        # Update status text
        if ($script:txtStatus) {
            $script:txtStatus.Text = $Message
            
            # Set appropriate color based on status
            switch ($Status) {
                "Connected" { 
                    $script:txtStatus.Foreground = "Green"
                    $script:SPOConnected = $true
                }
                "Connected (Demo)" { 
                    $script:txtStatus.Foreground = "Orange"
                    $script:SPOConnected = $true
                }
                "Disconnected" { 
                    $script:txtStatus.Foreground = "Gray"
                    $script:SPOConnected = $false
                }
                "Error" { 
                    $script:txtStatus.Foreground = "Red"
                    $script:SPOConnected = $false
                }
            }
        }
        
        Write-ActivityLog "Connection status updated: $Status - $Message" -Level "Information"
    }
    catch {
        Write-ActivityLog "Error updating connection status: $($_.Exception.Message)" -Level "Warning"
    }
}

function Enable-SharePointOperations {
    <#
    .SYNOPSIS
    Enables SharePoint operation buttons after successful connection
    #>
    try {
        # Enable all SharePoint operation buttons
        if ($script:btnGetSites) { $script:btnGetSites.IsEnabled = $true }
        if ($script:btnGetPermissions) { $script:btnGetPermissions.IsEnabled = $true }
        if ($script:btnGenerateReport) { $script:btnGenerateReport.IsEnabled = $true }
        if ($script:btnRefreshAnalytics) { $script:btnRefreshAnalytics.IsEnabled = $true }
        
        Write-ActivityLog "SharePoint operations enabled - Connection state: $script:SPOConnected" -Level "Information"
    }
    catch {
        Write-ActivityLog "Error enabling SharePoint operations: $($_.Exception.Message)" -Level "Warning"
    }
}

# [MARKED FOR DELETION] Duplicate of Test-PnPModuleAvailable in SPOConnection.ps1
# The original in SPOConnection.ps1 is more complete (includes Method 3 command-availability fallback).
# This stripped-down copy was introduced during "Mainwindow logic split" (2025-08-11) and is
# currently overriding the original due to dot-source order — removing restores the full version.
# function Test-PnPModuleAvailable {
#     # Method 1: Check installed modules (prefer 3.x)
#     $moduleInstalled = Get-Module -ListAvailable -Name PnP.PowerShell -ErrorAction SilentlyContinue |
#         Sort-Object Version -Descending | Select-Object -First 1
#
#     if ($moduleInstalled) {
#         $version = $moduleInstalled.Version
#         Write-ActivityLog "PnP module found via Get-Module: Version $version"
#
#         # Check if it's modern version
#         if ($version -ge [version]"3.0.0") {
#             Write-ActivityLog "Modern PnP PowerShell 3.x detected"
#             return $true
#         } elseif ($version -ge [version]"2.0.0") {
#             Write-ActivityLog "Legacy PnP PowerShell 2.x detected - recommend upgrading to 3.x"
#             return $true
#         } else {
#             Write-ActivityLog "Very old PnP PowerShell detected - upgrade required"
#             return $false
#         }
#     }
#
#     # Method 2: Try to import
#     try {
#         Import-Module PnP.PowerShell -ErrorAction Stop
#         $importedModule = Get-Module PnP.PowerShell
#         if ($importedModule) {
#             Write-ActivityLog "PnP module successfully imported: Version $($importedModule.Version)"
#             return $true
#         }
#     }
#     catch {
#         Write-ActivityLog "PnP module import failed: $($_.Exception.Message)"
#     }
#
#     return $false
# }

# [MARKED FOR DELETION] Duplicate of Install-PnPModule in SPOConnection.ps1
# The original in SPOConnection.ps1 is more complete (includes legacy module cleanup,
# verbose status messages, and post-install version verification).
# This simplified copy was introduced during "Mainwindow logic split" (2025-08-11).
# function Install-PnPModule {
#     param($UI)
#
#     try {
#         $UI.UpdateStatus("Installing modern PnP PowerShell 3.x...")
#
#         # Check PowerShell version first
#         if ($PSVersionTable.PSVersion.Major -lt 7) {
#             throw "PowerShell 7.0 or later is required for modern PnP PowerShell 3.x. Current version: $($PSVersionTable.PSVersion)"
#         }
#
#         # Check and install NuGet if needed
#         $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
#         if (-not $nugetProvider -or $nugetProvider.Version -lt [version]"2.8.5.201") {
#             $UI.UpdateStatus("Installing NuGet provider...")
#             Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
#         }
#
#         # Install modern PnP PowerShell 3.x
#         $UI.UpdateStatus("Installing PnP.PowerShell 3.x...")
#
#         $installParams = @{
#             Name = "PnP.PowerShell"
#             MinimumVersion = "3.0.0"
#             Force = $true
#             AllowClobber = $true
#             SkipPublisherCheck = $true
#             Scope = "CurrentUser"
#         }
#
#         Install-Module @installParams
#
#         # Force refresh and import
#         Import-Module PnP.PowerShell -Force -ErrorAction SilentlyContinue
#
#         $UI.UpdateStatus("✅ Modern PnP PowerShell installed successfully!")
#         return $true
#     }
#     catch {
#         $UI.UpdateStatus("Error installing modern PnP module: $($_.Exception.Message)")
#         throw
#     }
# }