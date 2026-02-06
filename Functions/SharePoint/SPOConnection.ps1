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
        $modernCommands = @("Connect-PnPOnline", "Get-PnPWeb", "Get-PnPRoleAssignment")
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

function Install-PnPModule {
    param($UI)

    try {
        $UI.UpdateStatus("Installing modern PnP PowerShell 3.x...`nThis may take a few minutes.")
        
        # Check PowerShell version first
        if ($PSVersionTable.PSVersion.Major -lt 7) {
            throw "PowerShell 7.0 or later is required for modern PnP PowerShell 3.x. Current version: $($PSVersionTable.PSVersion)"
        }
        
        # Remove any legacy versions first
        $UI.UpdateStatus("Cleaning up legacy PnP PowerShell versions...")
        $legacyModules = @("SharePointPnPPowerShellOnline", "PnP.PowerShell")
        
        foreach ($module in $legacyModules) {
            try {
                $installed = Get-Module -Name $module -ListAvailable -ErrorAction SilentlyContinue
                if ($installed) {
                    $oldVersions = $installed | Where-Object { $_.Version -lt [version]"3.0.0" }
                    if ($oldVersions) {
                        Write-ActivityLog "Removing legacy $module versions: $($oldVersions.Version -join ', ')"
                        Uninstall-Module -Name $module -Force -AllVersions -ErrorAction SilentlyContinue
                    }
                }
            }
            catch {
                Write-ActivityLog "Could not remove legacy $module`: $($_.Exception.Message)"
            }
        }
        
        # Check and install NuGet if needed
        $nugetProvider = Get-PackageProvider -Name NuGet -ErrorAction SilentlyContinue
        if (-not $nugetProvider -or $nugetProvider.Version -lt [version]"2.8.5.201") {
            $UI.UpdateStatus("Installing NuGet provider...")
            Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Scope CurrentUser | Out-Null
        }

        # Install modern PnP PowerShell 3.x
        $UI.UpdateStatus("Installing PnP.PowerShell 3.x...`nPlease wait, this may take several minutes.")
        
        $installParams = @{
            Name = "PnP.PowerShell"
            MinimumVersion = "3.0.0"
            Force = $true
            AllowClobber = $true
            SkipPublisherCheck = $true
            Scope = "CurrentUser"
        }
        
        Install-Module @installParams
        
        # Force refresh and import
        Import-Module PnP.PowerShell -Force -ErrorAction SilentlyContinue
        
        # Verify modern installation
        $installedModule = Get-Module -ListAvailable -Name PnP.PowerShell | 
            Sort-Object Version -Descending | Select-Object -First 1
            
        if (-not $installedModule) {
            throw "PnP PowerShell module installation verification failed"
        }
        
        if ($installedModule.Version -lt [version]"3.0.0") {
            throw "Failed to install modern PnP PowerShell 3.x. Got version $($installedModule.Version) instead."
        }
        
        # Test key modern commands
        $modernCommands = @("Connect-PnPOnline", "Get-PnPWeb", "Get-PnPRoleAssignment", "Get-PnPUser", "Get-PnPGroup")
        $missingCommands = @()
        
        foreach ($cmd in $modernCommands) {
            if (-not (Get-Command $cmd -ErrorAction SilentlyContinue)) {
                $missingCommands += $cmd
            }
        }
        
        if ($missingCommands.Count -gt 0) {
            Write-ActivityLog "Warning: Missing modern commands: $($missingCommands -join ', ')"
        }
        
        $UI.UpdateStatus("✅ Modern PnP PowerShell $($installedModule.Version) installed successfully!`nAll key commands available for enhanced SharePoint analysis.")
        return $true
    }
    catch {
        $UI.UpdateStatus("Error installing modern PnP module: $($_.Exception.Message)")
        throw
    }
}

# function Connect-SharePointWithAppRegistration {
#     param(
#         $TenantUrl, 
#         $ClientId, 
#         $UI
#     )
    
#     try {
#         Write-ActivityLog "Starting modern SharePoint connection to: $TenantUrl with App Registration: $ClientId"
        
#         # Clear any existing connections first
#         try {
#             Disconnect-PnPOnline -ErrorAction SilentlyContinue
#             Write-ActivityLog "Cleared any existing connections"
#         }
#         catch {
#             # Ignore errors during disconnect
#         }
        
#         # Ensure modern module is imported
#         $UI.UpdateStatus("Importing modern PnP PowerShell 3.x...")
#         try {
#             Import-Module PnP.PowerShell -Force -ErrorAction Stop
#             $pnpModule = Get-Module PnP.PowerShell
#             Write-ActivityLog "Modern PnP PowerShell imported successfully: Version $($pnpModule.Version)"
            
#             if ($pnpModule.Version -lt [version]"3.0.0") {
#                 throw "Legacy PnP PowerShell detected. Please upgrade to 3.x for full functionality."
#             }
#         }
#         catch {
#             throw "Failed to import modern PnP PowerShell module: $($_.Exception.Message)"
#         }
        
#         # Modern connection approach - simplified for PnP 3.x
#         $UI.UpdateStatus("Connecting to SharePoint Online...`nUsing App Registration: $ClientId`nPlease complete authentication in the popup window.")
#         Write-ActivityLog "Attempting modern connection to $TenantUrl with ClientId $ClientId"
        
#         try {
#             # Modern PnP 3.x connection - more reliable
#             Connect-PnPOnline -Url $TenantUrl -ClientId $ClientId -Interactive
#             Write-ActivityLog "Modern connection command completed successfully"
#         }
#         catch {
#             throw "Failed to connect to SharePoint with app registration: $($_.Exception.Message)"
#         }
        
#         # Enhanced verification using modern PnP features
#         $UI.UpdateStatus("Verifying SharePoint connection...")
#         try {
#             # Modern verification approach
#             $context = Get-PnPContext -ErrorAction SilentlyContinue
#             $connection = Get-PnPConnection -ErrorAction SilentlyContinue
            
#             if ($null -eq $context -and $null -eq $connection) {
#                 throw "No SharePoint connection available after authentication"
#             }
            
#             Write-ActivityLog "SharePoint context/connection verified successfully"
            
#             # Get enhanced site information using modern PnP
#             $site = $null
#             $web = $null
#             try {
#                 # Try modern approach first
#                 $web = Get-PnPWeb -ErrorAction SilentlyContinue
#                 if ($web) {
#                     $site = @{ 
#                         Url = $web.Url; 
#                         Title = $web.Title;
#                         Created = $web.Created;
#                         Language = $web.Language;
#                         Template = $web.WebTemplate
#                     }
#                     Write-ActivityLog "Enhanced site verification successful: $($web.Url)"
#                 } else {
#                     # Fallback
#                     $site = @{ Url = $TenantUrl; Title = "SharePoint Site" }
#                     Write-ActivityLog "Using basic verification - connection established"
#                 }
#             }
#             catch {
#                 Write-ActivityLog "Site verification had issues but proceeding: $($_.Exception.Message)"
#                 $site = @{ Url = $TenantUrl; Title = "SharePoint Site" }
#             }
            
#         }
#         catch {
#             Write-ActivityLog "Connection verification had issues but proceeding: $($_.Exception.Message)"
#             $site = @{ Url = $TenantUrl; Title = "SharePoint Site" }
#         }
        
#         $script:SPOConnected = $true
#         $script:SPOContext = Get-PnPConnection -ErrorAction SilentlyContinue
        
#         # Enhanced user info with modern PnP
#         $currentUser = ""
#         try {
#             $userInfo = Get-PnPCurrentUser -ErrorAction SilentlyContinue
#             if ($userInfo) {
#                 $currentUser = $userInfo.UserPrincipalName -or $userInfo.Email -or $userInfo.LoginName -or $userInfo.Title
#             }
#         }
#         catch {
#             $currentUser = "Authentication verified"
#         }
        
#         Write-ActivityLog "Successfully connected to SharePoint: $($site.Url) using modern PnP with app $ClientId"
#         $UI.UpdateStatus("✅ Successfully connected to SharePoint Online!`nSite: $($site.Url)`nApp Registration: $ClientId`nUser: $currentUser`nPnP Version: Modern 3.x`nReady for enhanced SharePoint operations.")
        
#         return $true
#     }
#     catch {
#         $script:SPOConnected = $false
#         $script:SPOContext = $null
#         Write-ErrorLog -Message $_.Exception.Message -Location "Modern-SharePoint-AppRegistration-Connection"
#         $UI.UpdateStatus("SharePoint connection failed: $($_.Exception.Message)")
#         throw
#     }
# }

# [MARKED FOR DELETION] Duplicate of Get-DetailedSiteInformation in SitesDeepDive.ps1
# The version in SitesDeepDive.ps1 is more complete (admin URL escalation, per-site storage
# retrieval). This earlier draft was introduced in the same commit ("init deepdive", 2025-08-19)
# and is already overridden at runtime since SitesDeepDive.ps1 loads last.
# function Get-DetailedSiteInformation {
#     <#
#     .SYNOPSIS
#     Gets detailed site information for deep dive analysis
#     #>
#     try {
#         Write-ActivityLog "Fetching detailed site information for deep dive" -Level "Information"
#
#         $detailedSites = @()
#
#         if ($script:SPOConnected) {
#             # Try to get tenant-level sites first
#             try {
#                 $sites = Get-PnPTenantSite -ErrorAction Stop
#
#                 foreach ($site in $sites) {
#                     $siteData = @{
#                         Title = $site.Title
#                         Url = $site.Url
#                         Owner = if ($site.Owner) { $site.Owner } else { $site.SiteOwnerEmail }
#                         Storage = if ($site.StorageUsageCurrent) { $site.StorageUsageCurrent.ToString() } else { "0" }
#                         StorageQuota = if ($site.StorageQuota) { $site.StorageQuota } else { 0 }
#                         Template = if ($site.Template) { $site.Template } else { "N/A" }
#                         LastModified = if ($site.LastContentModifiedDate) { $site.LastContentModifiedDate.ToString("yyyy-MM-dd") } else { "N/A" }
#                         Created = if ($site.Created) { $site.Created.ToString("yyyy-MM-dd") } else { "N/A" }
#                         IsHubSite = if ($site.IsHubSite) { $site.IsHubSite } else { $false }
#                         HubSiteId = if ($site.HubSiteId) { $site.HubSiteId } else { $null }
#                         SharingCapability = if ($site.SharingCapability) { $site.SharingCapability } else { "N/A" }
#                         Status = if ($site.Status) { $site.Status } else { "Active" }
#                         LockState = if ($site.LockState) { $site.LockState } else { "Unlock" }
#                         WebsCount = if ($site.WebsCount) { $site.WebsCount } else { 0 }
#                         HasUniquePermissions = $false  # Will be updated if we can connect to the site
#                     }
#
#                     # Try to get additional details by connecting to each site
#                     try {
#                         Connect-PnPOnline -Url $site.Url -ClientId (Get-AppSetting -SettingName "SharePoint.ClientId") -Interactive -ErrorAction SilentlyContinue
#                         $web = Get-PnPWeb -ErrorAction SilentlyContinue
#
#                         if ($web) {
#                             $siteData["HasUniquePermissions"] = $web.HasUniqueRoleAssignments
#
#                             # Get user and group counts if possible
#                             try {
#                                 $users = Get-PnPUser -ErrorAction SilentlyContinue
#                                 $groups = Get-PnPGroup -ErrorAction SilentlyContinue
#
#                                 if ($users) { $siteData["UserCount"] = $users.Count }
#                                 if ($groups) { $siteData["GroupCount"] = $groups.Count }
#                             }
#                             catch { }
#                         }
#                     }
#                     catch {
#                         Write-ActivityLog "Could not get additional details for site: $($site.Url)" -Level "Warning"
#                     }
#
#                     $detailedSites += $siteData
#                 }
#             }
#             catch {
#                 Write-ActivityLog "Tenant-level enumeration failed, using fallback" -Level "Warning"
#
#                 # Fallback: Get current site details only
#                 try {
#                     $web = Get-PnPWeb -ErrorAction Stop
#                     $site = Get-PnPSite -ErrorAction Stop
#
#                     $siteData = @{
#                         Title = $web.Title
#                         Url = $web.Url
#                         Owner = "Current User"
#                         Storage = if ($site.Usage) { [math]::Round($site.Usage.Storage / 1MB, 0).ToString() } else { "0" }
#                         StorageQuota = if ($site.Usage) { [math]::Round($site.Usage.StoragePercentageUsed, 0) } else { 0 }
#                         Template = $web.WebTemplate
#                         LastModified = if ($web.LastItemModifiedDate) { $web.LastItemModifiedDate.ToString("yyyy-MM-dd") } else { "N/A" }
#                         Created = if ($web.Created) { $web.Created.ToString("yyyy-MM-dd") } else { "N/A" }
#                         IsHubSite = $false
#                         HasUniquePermissions = $web.HasUniqueRoleAssignments
#                         Status = "Active"
#                         LockState = "Unlock"
#                     }
#
#                     $detailedSites += $siteData
#                 }
#                 catch {
#                     Write-ActivityLog "Failed to get site details: $($_.Exception.Message)" -Level "Error"
#                 }
#             }
#         }
#
#         Write-ActivityLog "Retrieved detailed information for $($detailedSites.Count) sites" -Level "Information"
#         return $detailedSites
#
#     }
#     catch {
#         Write-ErrorLog -Message $_.Exception.Message -Location "Get-DetailedSiteInformation"
#         return @()
#     }
# }

# [MARKED FOR DELETION] Duplicate of Get-SiteHealthMetrics in SitesDeepDive.ps1
# Both copies are character-identical, introduced in the same commit ("init deepdive", 2025-08-19).
# The canonical home is SitesDeepDive.ps1 which already overrides this at runtime.
# function Get-SiteHealthMetrics {
#     <#
#     .SYNOPSIS
#     Gets health metrics for a specific site
#     #>
#     param(
#         [Parameter(Mandatory=$true)]
#         [string]$SiteUrl
#     )
#
#     try {
#         Write-ActivityLog "Getting health metrics for site: $SiteUrl" -Level "Information"
#
#         $healthMetrics = @{
#             SiteUrl = $SiteUrl
#             HealthScore = 100  # Start with perfect score
#             Issues = @()
#             Recommendations = @()
#         }
#
#         # Connect to the site
#         Connect-PnPOnline -Url $SiteUrl -ClientId (Get-AppSetting -SettingName "SharePoint.ClientId") -Interactive
#
#         # Check storage usage
#         $site = Get-PnPSite
#         if ($site.Usage) {
#             $storagePercentage = $site.Usage.StoragePercentageUsed
#
#             if ($storagePercentage -gt 90) {
#                 $healthMetrics.HealthScore -= 30
#                 $healthMetrics.Issues += "Critical: Storage usage above 90%"
#                 $healthMetrics.Recommendations += "Immediately archive or delete unnecessary content"
#             }
#             elseif ($storagePercentage -gt 75) {
#                 $healthMetrics.HealthScore -= 15
#                 $healthMetrics.Issues += "Warning: Storage usage above 75%"
#                 $healthMetrics.Recommendations += "Plan for storage cleanup or quota increase"
#             }
#         }
#
#         # Check for orphaned permissions
#         $web = Get-PnPWeb
#         if ($web.HasUniqueRoleAssignments) {
#             $roleAssignments = Get-PnPProperty -ClientObject $web -Property RoleAssignments
#
#             # Check for empty groups or invalid principals
#             foreach ($assignment in $roleAssignments) {
#                 try {
#                     $member = Get-PnPProperty -ClientObject $assignment -Property Member
#                     if (-not $member -or $member.PrincipalType -eq "None") {
#                         $healthMetrics.HealthScore -= 10
#                         $healthMetrics.Issues += "Orphaned permission assignment detected"
#                         $healthMetrics.Recommendations += "Review and clean up permission assignments"
#                         break
#                     }
#                 }
#                 catch { }
#             }
#         }
#
#         # Check for external sharing
#         $lists = Get-PnPList
#         $externalSharingDetected = $false
#
#         foreach ($list in $lists | Where-Object { -not $_.Hidden }) {
#             if ($list.HasUniqueRoleAssignments) {
#                 $healthMetrics.HealthScore -= 5
#                 $externalSharingDetected = $true
#                 break
#             }
#         }
#
#         if ($externalSharingDetected) {
#             $healthMetrics.Issues += "Multiple lists with unique permissions detected"
#             $healthMetrics.Recommendations += "Review list-level permissions for consistency"
#         }
#
#         # Check last activity
#         if ($web.LastItemModifiedDate) {
#             $daysSinceModified = (Get-Date) - $web.LastItemModifiedDate
#
#             if ($daysSinceModified.Days -gt 180) {
#                 $healthMetrics.HealthScore -= 20
#                 $healthMetrics.Issues += "Site inactive for over 180 days"
#                 $healthMetrics.Recommendations += "Consider archiving or deleting if no longer needed"
#             }
#             elseif ($daysSinceModified.Days -gt 90) {
#                 $healthMetrics.HealthScore -= 10
#                 $healthMetrics.Issues += "Site inactive for over 90 days"
#                 $healthMetrics.Recommendations += "Review if site is still actively needed"
#             }
#         }
#
#         # Ensure score doesn't go below 0
#         if ($healthMetrics.HealthScore -lt 0) { $healthMetrics.HealthScore = 0 }
#
#         Write-ActivityLog "Health metrics calculated for site: Score = $($healthMetrics.HealthScore)" -Level "Information"
#         return $healthMetrics
#
#     }
#     catch {
#         Write-ErrorLog -Message $_.Exception.Message -Location "Get-SiteHealthMetrics"
#         return $null
#     }
# }
