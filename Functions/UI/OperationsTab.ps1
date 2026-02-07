function Initialize-OperationsTab {
    <#
    .SYNOPSIS
    Initializes the SharePoint Operations tab with event handlers
    #>
    try {
        Write-ActivityLog "Initializing Operations tab with data-driven approach" -Level "Information"

        # Initialize data manager
        Initialize-SharePointDataManager

        # Set up event handlers using data-driven approach
        $script:btnGetSites.Add_Click({
            if (Get-AppSetting -SettingName "DemoMode") {
                Get-DemoSites-DataDriven
            } else {
                Get-RealSites-DataDriven
            }
        })

        $script:btnGetPermissions.Add_Click({
            if (Get-AppSetting -SettingName "DemoMode") {
                Get-DemoPermissions-DataDriven
            } else {
                Get-RealPermissions-DataDriven
            }
        })

        $script:btnGenerateReport.Add_Click({
            if (Get-AppSetting -SettingName "DemoMode") {
                Generate-DemoReport
            } else {
                Generate-RealReport
            }
        })

        # Set initial state
        $script:txtOperationsResults.Text = "Connect to SharePoint to begin operations..."

        # Disable buttons initially
        $script:btnGetSites.IsEnabled = $false
        $script:btnGetPermissions.IsEnabled = $false
        $script:btnGenerateReport.IsEnabled = $false

        Write-ActivityLog "Operations tab initialized successfully with data-driven approach" -Level "Information"
    }
    catch {
        Write-ActivityLog "Failed to initialize Operations tab: $($_.Exception.Message)" -Level "Error"
        throw
    }
}

function Get-RealSites-DataDriven {
    <#
    .SYNOPSIS
    Retrieves real SharePoint sites with proper storage data
    #>
    try {
        if (-not $script:SPOConnected) {
            throw "Not connected to SharePoint. Please connect first."
        }

        # Clear previous sites data and set context
        Clear-SharePointData -DataType "Sites"
        Set-SharePointOperationContext -OperationType "Sites Analysis"
        Start-Checkpoint -OperationType "SitesAnalysis" -Scope "Tenant"
        Reset-ThrottleStats

        # Start audit session
        Start-AuditSession -OperationType "SitesAnalysis" -ScanScope "Tenant"

        # Clear console and show header
        Write-ConsoleOutput "🔍 SHAREPOINT SITES ANALYSIS" -ForceUpdate
        Write-ConsoleOutput "=====================================================" -Append -ForceUpdate
        Write-ConsoleOutput "⏱️ Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append -ForceUpdate

        Write-ConsoleOutput "🔄 Using modern PnP PowerShell for site enumeration..." -Append -ForceUpdate
        Write-ConsoleOutput "📡 Attempting tenant-level site enumeration..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 1000

        $sites = @()
        $requiresAdmin = $false

        # Try to get sites with full details
        try {
            Write-ConsoleOutput "🔍 Scanning tenant for site collections..." -Append -ForceUpdate

            # Check if we need admin connection
            $tenantUrl = Get-AppSetting -SettingName "SharePoint.TenantUrl"
            $adminUrl = $tenantUrl -replace "\.sharepoint\.com", "-admin.sharepoint.com"
            $clientId = Get-AppSetting -SettingName "SharePoint.ClientId"

            # Try admin connection first for best results
            try {
                Write-ConsoleOutput "🔐 Attempting admin center connection for full details..." -Append -ForceUpdate
                Connect-PnPOnline -Url $adminUrl -ClientId $clientId -Interactive -ErrorAction Stop
                $sites = Invoke-WithThrottleProtection -OperationName "Get-PnPTenantSite (admin)" -ScriptBlock {
                    Get-PnPTenantSite -Detailed -ErrorAction Stop
                }
                $requiresAdmin = $true
                Write-ActivityLog "Retrieved $($sites.Count) sites with full details from admin center"
                Write-ConsoleOutput "✅ Successfully retrieved $($sites.Count) sites with storage details" -Append -ForceUpdate
            }
            catch {
                Write-ActivityLog "Admin connection failed, trying regular connection: $($_.Exception.Message)"
                Write-ConsoleOutput "⚠️ Admin access unavailable, trying standard enumeration..." -Append -ForceUpdate

                # Fallback to regular tenant connection
                Connect-PnPOnline -Url $tenantUrl -ClientId $clientId -Interactive -ErrorAction Stop
                $sites = Invoke-WithThrottleProtection -OperationName "Get-PnPTenantSite" -ScriptBlock {
                    Get-PnPTenantSite -ErrorAction Stop
                }
                Write-ConsoleOutput "✅ Retrieved $($sites.Count) sites (limited details)" -Append -ForceUpdate
            }
        }
        catch {
            Write-ActivityLog "Tenant-level enumeration failed: $($_.Exception.Message)"
            Write-ConsoleOutput "⚠️ Tenant-level access limited, using fallback method..." -Append -ForceUpdate

            # Final fallback - get current site only
            try {
                $currentWeb = Get-PnPWeb -ErrorAction Stop
                $currentSite = Get-PnPSite -Includes Usage -ErrorAction Stop

                # Create a site object that matches the structure
                $siteObj = [PSCustomObject]@{
                    Title = $currentWeb.Title
                    Url = $currentWeb.Url
                    StorageUsageCurrent = if ($currentSite.Usage -and $currentSite.Usage.Storage) {
                        [math]::Round($currentSite.Usage.Storage / 1MB, 0)
                    } else { 0 }
                    StorageQuota = if ($currentSite.Usage -and $currentSite.Usage.StoragePercentageUsed) {
                        [math]::Round($currentSite.Usage.StoragePercentageUsed * 100, 0)
                    } else { 0 }
                    Template = $currentWeb.WebTemplate
                    LastContentModifiedDate = $currentWeb.LastItemModifiedDate
                    Owner = "Current User"
                }

                $sites = @($siteObj)
                Write-ConsoleOutput "📂 Found current site only" -Append -ForceUpdate
            }
            catch {
                throw "Unable to retrieve any sites. SharePoint Administrator permissions may be required."
            }
        }

        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "📊 SITES DISCOVERY RESULTS" -Append -ForceUpdate
        Write-ConsoleOutput "=" * 50 -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "🏢 Sites Found: $($sites.Count)" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append

        # Process and store each site with proper storage data
        $siteCounter = 0
        foreach ($site in $sites | Select-Object -First 25) {
            $siteCounter++

            # Extract storage value properly based on the object type
            $storageValue = 0

            # For tenant sites (from Get-PnPTenantSite)
            if ($site.StorageUsageCurrent -ne $null) {
                $storageValue = [int]$site.StorageUsageCurrent
            }
            # For site objects with Usage property
            elseif ($site.Usage -and $site.Usage.Storage) {
                $storageValue = [math]::Round($site.Usage.Storage / 1MB, 0)
            }
            # Try to get storage by connecting to the site
            elseif ($site.Url) {
                try {
                    Connect-PnPOnline -Url $site.Url -ClientId $clientId -Interactive -ErrorAction SilentlyContinue
                    $siteDetail = Get-PnPSite -Includes Usage -ErrorAction SilentlyContinue
                    if ($siteDetail -and $siteDetail.Usage -and $siteDetail.Usage.Storage) {
                        $storageValue = [math]::Round($siteDetail.Usage.Storage / 1MB, 0)
                    }
                }
                catch {
                    Write-ActivityLog "Could not get storage for site: $($site.Url)" -Level "Warning"
                }
            }

            # Create site data object with proper storage
            $siteData = @{
                Title = if ($site.Title) { $site.Title } else { "Site $siteCounter" }
                Url = if ($site.Url) { $site.Url } else { "N/A" }
                Owner = if ($site.Owner) { $site.Owner } elseif ($site.SiteOwnerEmail) { $site.SiteOwnerEmail } else { "N/A" }
                Storage = $storageValue.ToString()  # Real storage value
                StorageQuota = if ($site.StorageQuota) { $site.StorageQuota } else { 0 }
                Template = if ($site.Template) { $site.Template } else { "N/A" }
                LastModified = if ($site.LastContentModifiedDate) {
                    $site.LastContentModifiedDate.ToString("yyyy-MM-dd")
                } else { "N/A" }
                IsHubSite = if ($site.IsHubSite) { $site.IsHubSite } else { $false }
                UserCount = 0
                GroupCount = 0
            }

            # Add site to data store
            Add-SharePointSite -SiteData $siteData

            # Display in console with real storage value
            Write-ConsoleOutput "📁 SITE #$siteCounter`: $($siteData.Title)" -Append -ForceUpdate
            Write-ConsoleOutput "   🌐 URL: $($siteData.Url)" -Append
            Write-ConsoleOutput "   👤 Owner: $($siteData.Owner)" -Append
            Write-ConsoleOutput "   💾 Storage: $storageValue MB" -Append  # Show real value
            if ($siteData.StorageQuota -gt 0) {
                $quotaMB = [math]::Round($siteData.StorageQuota / 1MB, 0)
                $percentUsed = if ($quotaMB -gt 0) {
                    [math]::Round(($storageValue / $quotaMB) * 100, 1)
                } else { 0 }
                Write-ConsoleOutput "   📊 Storage Quota: $quotaMB MB ($percentUsed% used)" -Append
            }
            if ($siteData.Template -ne "N/A") { Write-ConsoleOutput "   🎨 Template: $($siteData.Template)" -Append }
            if ($siteData.LastModified -ne "N/A") { Write-ConsoleOutput "   📅 Last Modified: $($siteData.LastModified)" -Append }
            if ($siteData.IsHubSite) { Write-ConsoleOutput "   🌟 Hub Site" -Append }
            Write-ConsoleOutput "   ============================================================" -Append
            Update-UIAndWait -WaitMs 300
        }

        if ($sites.Count -gt 25) {
            Write-ConsoleOutput "... and $($sites.Count - 25) more sites" -Append
            Write-ConsoleOutput "" -Append
        }

        # Show storage summary
        $totalStorage = 0
        $sitesWithStorage = 0
        $allSites = Get-SharePointData -DataType "Sites"
        foreach ($s in $allSites) {
            $storage = [int]$s["Storage"]
            if ($storage -gt 0) {
                $totalStorage += $storage
                $sitesWithStorage++
            }
        }

        if ($sitesWithStorage -gt 0) {
            $totalGB = [math]::Round($totalStorage / 1024, 2)
            $avgMB = [math]::Round($totalStorage / $sitesWithStorage, 0)
            Write-ConsoleOutput "" -Append
            Write-ConsoleOutput "💾 STORAGE SUMMARY:" -Append -ForceUpdate
            Write-ConsoleOutput "   • Total Storage Used: $totalStorage MB ($totalGB GB)" -Append
            Write-ConsoleOutput "   • Average per Site: $avgMB MB" -Append
            Write-ConsoleOutput "   • Sites with Storage Data: $sitesWithStorage/$($sites.Count)" -Append
        }

        if (-not $requiresAdmin -and $sitesWithStorage -eq 0) {
            Write-ConsoleOutput "" -Append
            Write-ConsoleOutput "⚠️ Note: Storage data requires SharePoint Administrator permissions" -Append
            Write-ConsoleOutput "   To get storage details, please ensure:" -Append
            Write-ConsoleOutput "   1. You have SharePoint Administrator role" -Append
            Write-ConsoleOutput "   2. Or connect to the admin center URL" -Append
        }

        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ Site enumeration completed successfully!" -Append -ForceUpdate

        # UPDATE VISUAL ANALYTICS DIRECTLY FROM DATA
        Update-VisualAnalyticsFromData

        Write-ActivityLog "Sites operation completed with $($sites.Count) sites" -Level "Information"
        Complete-Checkpoint -Status "Completed"
        Write-AuditEvent -EventType "DataCollection" -Detail "Sites analysis complete: $($sites.Count) sites"
        Complete-AuditSession -Status "Completed"

        # Log throttle stats if any retries occurred
        $throttleStats = Get-ThrottleStats
        if ($throttleStats.TotalRetries -gt 0) {
            Write-ConsoleOutput "⚠️ Throttle protection: $($throttleStats.TotalRetries) retries, $($throttleStats.TotalThrottled) throttle events" -Append
        }

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-RealSites-DataDriven"
        Complete-Checkpoint -Status "Failed"
        Write-AuditEvent -EventType "Error" -Detail $_.Exception.Message
        Complete-AuditSession -Status "Failed"
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "❌ ERROR: Site enumeration failed" -Append
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
    }
}

function Get-RealPermissions-DataDriven {
    <#
    .SYNOPSIS
    Analyzes real SharePoint permissions and stores in data manager with REAL storage data
    #>
    try {
        if (-not $script:SPOConnected) {
            throw "Not connected to SharePoint. Please connect first."
        }

        $siteUrl = $script:txtSiteUrl.Text.Trim()
        if ([string]::IsNullOrEmpty($siteUrl)) {
            Write-ConsoleOutput "Please enter a specific site URL to analyze permissions." -Append
            return
        }

        # Clear previous data and set context
        Clear-SharePointData -DataType "All"
        Set-SharePointOperationContext -OperationType "Permissions Analysis"
        Start-Checkpoint -OperationType "PermissionsAnalysis" -Scope $siteUrl
        Reset-ThrottleStats

        # Start audit session
        Start-AuditSession -OperationType "PermissionsAnalysis" -ScanScope $siteUrl

        # Clear console and show header
        Write-ConsoleOutput "🔐 SHAREPOINT PERMISSIONS ANALYSIS" -ForceUpdate
        Write-ConsoleOutput "=====================================================" -Append -ForceUpdate
        Write-ConsoleOutput "⏱️ Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append -ForceUpdate
        Write-ConsoleOutput "🎯 Target: $siteUrl" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append -ForceUpdate

        Write-ConsoleOutput "🔄 Analyzing permissions for: $siteUrl..." -Append -ForceUpdate
        Write-ConsoleOutput "📡 Connecting to the specified site..." -Append -ForceUpdate

        # Connect to the site
        Connect-PnPOnline -Url $siteUrl -ClientId (Get-AppSetting -SettingName "SharePoint.ClientId") -Interactive
        Write-ConsoleOutput "✅ Connected successfully!" -Append -ForceUpdate

        # Get site information INCLUDING STORAGE
        Write-ConsoleOutput "📊 Retrieving site information and storage data..." -Append -ForceUpdate

        $web = Get-PnPWeb -ErrorAction Stop

        # Get REAL storage data
        $storageValue = 0
        $storageQuota = 0
        $storagePercentUsed = 0

        try {
            # Get site with Usage information
            $site = Get-PnPSite -Includes Usage -ErrorAction Stop

            if ($site.Usage) {
                # Storage is in bytes, convert to MB
                $storageValue = [math]::Round($site.Usage.Storage / 1MB, 0)
                $storageQuota = [math]::Round($site.Usage.StorageQuotaInMB, 0)

                if ($site.Usage.StoragePercentageUsed) {
                    $storagePercentUsed = [math]::Round($site.Usage.StoragePercentageUsed, 2)
                }

                Write-ActivityLog "Retrieved storage data: $storageValue MB used of $storageQuota MB quota ($storagePercentUsed%)" -Level "Information"
            }
            else {
                Write-ActivityLog "Site.Usage property is null" -Level "Warning"
            }
        }
        catch {
            Write-ActivityLog "Failed to get storage data: $($_.Exception.Message)" -Level "Warning"

            # Try alternative method using tenant admin if available
            try {
                $tenantUrl = Get-AppSetting -SettingName "SharePoint.TenantUrl"
                $adminUrl = $tenantUrl -replace "\.sharepoint\.com", "-admin.sharepoint.com"

                Write-ConsoleOutput "⚠️ Attempting to get storage via admin connection..." -Append -ForceUpdate

                Connect-PnPOnline -Url $adminUrl -ClientId (Get-AppSetting -SettingName "SharePoint.ClientId") -Interactive -ErrorAction Stop
                $tenantSites = Get-PnPTenantSite -Url $siteUrl -Detailed -ErrorAction Stop

                if ($tenantSites -and $tenantSites.StorageUsageCurrent) {
                    $storageValue = [int]$tenantSites.StorageUsageCurrent
                    $storageQuota = if ($tenantSites.StorageQuota) { [int]$tenantSites.StorageQuota } else { 0 }
                    Write-ActivityLog "Retrieved storage from admin: $storageValue MB" -Level "Information"
                }

                # Reconnect to the original site
                Connect-PnPOnline -Url $siteUrl -ClientId (Get-AppSetting -SettingName "SharePoint.ClientId") -Interactive
            }
            catch {
                Write-ActivityLog "Admin storage retrieval also failed: $($_.Exception.Message)" -Level "Warning"
            }
        }

        # Add the analyzed site to data store with REAL storage
        $siteData = @{
            Title = if ($web.Title) { $web.Title } else { "Analyzed Site" }
            Url = if ($web.Url) { $web.Url } else { $siteUrl }
            Owner = "Current User"
            Storage = $storageValue.ToString()  # REAL storage value, not hardcoded!
            StorageQuota = $storageQuota.ToString()
            Template = if ($web.WebTemplate) { $web.WebTemplate } else { "N/A" }
            HasUniquePermissions = $web.HasUniqueRoleAssignments
            Created = if ($web.Created) { $web.Created.ToString("yyyy-MM-dd") } else { "N/A" }
            LastModified = if ($web.LastItemModifiedDate) { $web.LastItemModifiedDate.ToString("yyyy-MM-dd") } else { "N/A" }
        }
        Add-SharePointSite -SiteData $siteData

        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ SITE INFORMATION" -Append -ForceUpdate
        Write-ConsoleOutput "📝 Title: $($web.Title)" -Append
        Write-ConsoleOutput "🌐 URL: $($web.Url)" -Append
        Write-ConsoleOutput "💾 Storage Used: $storageValue MB" -Append  # Show REAL value
        if ($storageQuota -gt 0) {
            Write-ConsoleOutput "📊 Storage Quota: $storageQuota MB ($storagePercentUsed% used)" -Append
        }
        Write-ConsoleOutput "🎨 Template: $($web.WebTemplate)" -Append
        Write-ConsoleOutput "📅 Created: $($siteData.Created)" -Append
        Write-ConsoleOutput "📅 Last Modified: $($siteData.LastModified)" -Append
        Write-ConsoleOutput "🔒 Has Unique Permissions: $($web.HasUniqueRoleAssignments)" -Append
        Write-ConsoleOutput "" -Append

        # Get and store users
        Write-ConsoleOutput "👥 Retrieving users..." -Append -ForceUpdate
        Update-Checkpoint -Phase "Users"
        try {
            $users = Invoke-WithThrottleProtection -OperationName "Get-PnPUser" -ScriptBlock {
                Get-PnPUser -ErrorAction Stop
            }
            $userCounter = 0
            $regularUsers = $users | Where-Object {
                $_.PrincipalType -eq "User" -and
                -not $_.LoginName.Contains("app@sharepoint") -and
                -not $_.LoginName.Contains("SHAREPOINT\system")
            }

            foreach ($user in $regularUsers) {
                $userData = @{
                    Name = if ($user.Title) { $user.Title } else { "Unknown User" }
                    Email = if ($user.Email) { $user.Email } else { "N/A" }
                    LoginName = $user.LoginName
                    Type = if ($user.IsShareByEmailGuestUser -or $user.IsEmailAuthenticationGuestUser) { "External" } else { "Internal" }
                    IsSiteAdmin = $user.IsSiteAdmin
                    Permission = if ($user.IsSiteAdmin) { "Full Control" } else { "Member" }
                }
                Add-SharePointUser -UserData $userData

                $userCounter++
                if ($userCounter -le 10) {
                    Write-ConsoleOutput "👤 USER #$userCounter`: $($userData.Name)" -Append
                    if ($userData.Email -ne "N/A") { Write-ConsoleOutput "   📧 Email: $($userData.Email)" -Append }
                    if ($userData.Type -eq "External") { Write-ConsoleOutput "   🌐 External User" -Append }
                    if ($userData.IsSiteAdmin) { Write-ConsoleOutput "   🔑 Site Administrator" -Append }
                    Write-ConsoleOutput "   --------------------------------------------------------" -Append
                    Update-UIAndWait -WaitMs 200
                }
            }

            if ($regularUsers.Count -gt 10) {
                Write-ConsoleOutput "   ... and $($regularUsers.Count - 10) more users" -Append
            }

            Write-ConsoleOutput "" -Append
            Write-ConsoleOutput "✅ Retrieved $($regularUsers.Count) users" -Append
            Write-ConsoleOutput "" -Append
        }
        catch {
            Write-ConsoleOutput "⚠️ Limited access to user information" -Append
            Write-ConsoleOutput "" -Append
        }

        # Get and store groups
        Write-ConsoleOutput "👨‍👩‍👧‍👦 Retrieving groups..." -Append -ForceUpdate
        Update-Checkpoint -Phase "Groups"
        try {
            $groups = Invoke-WithThrottleProtection -OperationName "Get-PnPGroup" -ScriptBlock {
                Get-PnPGroup -ErrorAction Stop
            }
            $groupCounter = 0

            $importantGroups = $groups | Where-Object {
                -not $_.Title.StartsWith("SharingLinks") -and
                -not $_.Title.StartsWith("Limited Access")
            }

            foreach ($group in $importantGroups) {
                $memberCount = 0
                try {
                    $members = Get-PnPGroupMember -Group $group.Title -ErrorAction SilentlyContinue
                    if ($members) { $memberCount = $members.Count }
                }
                catch { }

                $groupData = @{
                    Name = $group.Title
                    Description = if ($group.Description) { $group.Description } else { "N/A" }
                    MemberCount = $memberCount
                    Permission = "Group Permission"
                    Id = $group.Id
                }
                Add-SharePointGroup -GroupData $groupData

                $groupCounter++
                if ($groupCounter -le 10) {
                    Write-ConsoleOutput "👥 GROUP #$groupCounter`: $($groupData.Name)" -Append
                    if ($groupData.Description -ne "N/A") { Write-ConsoleOutput "   📝 Description: $($groupData.Description)" -Append }
                    Write-ConsoleOutput "   👥 Members: $memberCount" -Append
                    Write-ConsoleOutput "   ========================================================" -Append
                    Update-UIAndWait -WaitMs 250
                }
            }

            if ($importantGroups.Count -gt 10) {
                Write-ConsoleOutput "   ... and $($importantGroups.Count - 10) more groups" -Append
            }

            Write-ConsoleOutput "" -Append
            Write-ConsoleOutput "✅ Retrieved $($importantGroups.Count) groups" -Append
            Write-ConsoleOutput "" -Append
        }
        catch {
            Write-ConsoleOutput "❌ Failed to retrieve groups: $($_.Exception.Message)" -Append
            Write-ConsoleOutput "" -Append
        }

        # ===== ROLE ASSIGNMENT MAPPING =====
        Write-ConsoleOutput "🔑 Analyzing role assignments..." -Append -ForceUpdate
        Update-Checkpoint -Phase "RoleAssignments"
        try {
            $web = Get-PnPWeb -ErrorAction Stop
            $roleAssignments = Invoke-WithThrottleProtection -OperationName "Get RoleAssignments" -ScriptBlock {
                Get-PnPProperty -ClientObject $web -Property RoleAssignments -ErrorAction Stop
            }
            $raCounter = 0

            foreach ($ra in $roleAssignments) {
                try {
                    $member = Get-PnPProperty -ClientObject $ra -Property Member -ErrorAction SilentlyContinue
                    $roleBindings = Get-PnPProperty -ClientObject $ra -Property RoleDefinitionBindings -ErrorAction SilentlyContinue

                    if ($member -and $roleBindings) {
                        foreach ($roleDef in $roleBindings) {
                            if ($roleDef.Name -eq "Limited Access") { continue }

                            $principalType = switch ($member.PrincipalType) {
                                "User" { "User" }
                                "SharePointGroup" { "SharePoint Group" }
                                "SecurityGroup" { "Security Group" }
                                default { $member.PrincipalType.ToString() }
                            }

                            $roleData = @{
                                Principal = $member.Title
                                PrincipalType = $principalType
                                Role = $roleDef.Name
                                Scope = "Site"
                                ScopeUrl = $siteUrl
                                SiteTitle = $web.Title
                            }
                            Add-SharePointRoleAssignment -RoleData $roleData
                            $raCounter++
                        }
                    }
                }
                catch { }
            }

            Write-ConsoleOutput "✅ Found $raCounter site-level role assignments" -Append -ForceUpdate
            Write-ConsoleOutput "" -Append
        }
        catch {
            Write-ConsoleOutput "⚠️ Limited access to role assignment data: $($_.Exception.Message)" -Append
            Write-ConsoleOutput "" -Append
        }

        # ===== PERMISSION INHERITANCE TREE =====
        Write-ConsoleOutput "🔗 Checking permission inheritance..." -Append -ForceUpdate
        Update-Checkpoint -Phase "Inheritance"
        try {
            # Add site-level entry
            Add-SharePointInheritanceItem -InheritanceData @{
                Title = $web.Title
                Type = "Site"
                Url = $web.Url
                HasUniquePermissions = $web.HasUniqueRoleAssignments
                ParentUrl = "N/A"
                RoleAssignmentCount = $raCounter
                SiteTitle = $web.Title
            }

            $lists = Invoke-WithThrottleProtection -OperationName "Get-PnPList" -ScriptBlock {
                Get-PnPList -ErrorAction Stop
            }
            $visibleLists = $lists | Where-Object { -not $_.Hidden }
            $brokenCount = 0

            foreach ($list in $visibleLists) {
                $listType = if ($list.BaseType -eq "DocumentLibrary") { "Document Library" } else { "List" }

                $listRaCount = 0
                if ($list.HasUniqueRoleAssignments) {
                    $brokenCount++
                    try {
                        $listRAs = Get-PnPProperty -ClientObject $list -Property RoleAssignments -ErrorAction SilentlyContinue
                        if ($listRAs) { $listRaCount = $listRAs.Count }

                        # Also capture list-level role assignments
                        foreach ($listRA in $listRAs) {
                            try {
                                $listMember = Get-PnPProperty -ClientObject $listRA -Property Member -ErrorAction SilentlyContinue
                                $listRoleBindings = Get-PnPProperty -ClientObject $listRA -Property RoleDefinitionBindings -ErrorAction SilentlyContinue

                                if ($listMember -and $listRoleBindings) {
                                    foreach ($roleDef in $listRoleBindings) {
                                        if ($roleDef.Name -eq "Limited Access") { continue }

                                        $principalType = switch ($listMember.PrincipalType) {
                                            "User" { "User" }
                                            "SharePointGroup" { "SharePoint Group" }
                                            "SecurityGroup" { "Security Group" }
                                            default { $listMember.PrincipalType.ToString() }
                                        }

                                        Add-SharePointRoleAssignment -RoleData @{
                                            Principal = $listMember.Title
                                            PrincipalType = $principalType
                                            Role = $roleDef.Name
                                            Scope = $listType
                                            ScopeUrl = $list.RootFolder.ServerRelativeUrl
                                            SiteTitle = $web.Title
                                        }
                                    }
                                }
                            }
                            catch { }
                        }
                    }
                    catch { }
                }

                Add-SharePointInheritanceItem -InheritanceData @{
                    Title = $list.Title
                    Type = $listType
                    Url = $list.RootFolder.ServerRelativeUrl
                    HasUniquePermissions = $list.HasUniqueRoleAssignments
                    ParentUrl = $web.Url
                    RoleAssignmentCount = $listRaCount
                    SiteTitle = $web.Title
                }
            }

            Write-ConsoleOutput "✅ Scanned $($visibleLists.Count) lists/libraries - $brokenCount with broken inheritance" -Append -ForceUpdate
            Write-ConsoleOutput "" -Append
        }
        catch {
            Write-ConsoleOutput "⚠️ Limited access to inheritance data: $($_.Exception.Message)" -Append
            Write-ConsoleOutput "" -Append
        }

        # ===== SHARING LINKS AUDIT =====
        Write-ConsoleOutput "🔗 Auditing sharing links..." -Append -ForceUpdate
        Update-Checkpoint -Phase "SharingLinks"
        try {
            $allGroups = Invoke-WithThrottleProtection -OperationName "Get-PnPGroup (sharing)" -ScriptBlock {
                Get-PnPGroup -ErrorAction Stop
            }
            $sharingGroups = $allGroups | Where-Object { $_.Title.StartsWith("SharingLinks") }
            $linkCounter = 0

            foreach ($sg in $sharingGroups) {
                # Parse link type from group name
                $linkType = "Specific People"
                $accessLevel = "View"

                if ($sg.Title -match "AnonymousView") {
                    $linkType = "Anonymous"; $accessLevel = "View"
                } elseif ($sg.Title -match "AnonymousEdit") {
                    $linkType = "Anonymous"; $accessLevel = "Edit"
                } elseif ($sg.Title -match "OrganizationView") {
                    $linkType = "Company-wide"; $accessLevel = "View"
                } elseif ($sg.Title -match "OrganizationEdit") {
                    $linkType = "Company-wide"; $accessLevel = "Edit"
                } elseif ($sg.Title -match "Flexible") {
                    $linkType = "Specific People"; $accessLevel = "Edit"
                }

                $memberCount = 0
                try {
                    $sgMembers = Get-PnPGroupMember -Group $sg.Title -ErrorAction SilentlyContinue
                    if ($sgMembers) { $memberCount = $sgMembers.Count }
                }
                catch { }

                Add-SharePointSharingLink -LinkData @{
                    GroupName = $sg.Title
                    LinkType = $linkType
                    AccessLevel = $accessLevel
                    MemberCount = $memberCount
                    SiteTitle = $web.Title
                    CreatedDate = "N/A"
                    GroupId = $sg.Id
                }
                $linkCounter++
            }

            Write-ConsoleOutput "✅ Found $linkCounter sharing links" -Append -ForceUpdate
            Write-ConsoleOutput "" -Append
        }
        catch {
            Write-ConsoleOutput "⚠️ Limited access to sharing links data: $($_.Exception.Message)" -Append
            Write-ConsoleOutput "" -Append
        }

        # Show storage notice if we couldn't get it
        if ($storageValue -eq 0) {
            Write-ConsoleOutput "⚠️ Note: Storage data unavailable. This requires:" -Append
            Write-ConsoleOutput "   • SharePoint Administrator permissions, or" -Append
            Write-ConsoleOutput "   • Sites.FullControl.All API permission" -Append
            Write-ConsoleOutput "" -Append
        }

        # Show P3 analysis summary
        $p3Metrics = Get-SharePointData -DataType "Metrics"
        Write-ConsoleOutput "🔑 CORE ANALYSIS SUMMARY:" -Append -ForceUpdate
        Write-ConsoleOutput "   • Role Assignments: $($p3Metrics.TotalRoleAssignments)" -Append
        Write-ConsoleOutput "   • Inheritance Breaks: $($p3Metrics.InheritanceBreaks)" -Append
        Write-ConsoleOutput "   • Sharing Links: $($p3Metrics.TotalSharingLinks)" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "=" * 65 -Append
        Write-ConsoleOutput "✅ PERMISSIONS ANALYSIS COMPLETED SUCCESSFULLY" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "💡 TIP: Click metric tiles on Visual Analytics tab for deep dive analysis" -Append -ForceUpdate
        Write-ConsoleOutput "   New: Role Assignments, Inheritance Tree, and Sharing Links deep dives!" -Append
        Write-ConsoleOutput "" -Append

        # UPDATE VISUAL ANALYTICS DIRECTLY FROM DATA
        Update-VisualAnalyticsFromData

        Write-ActivityLog "Permissions analysis completed with storage: $storageValue MB" -Level "Information"
        Complete-Checkpoint -Status "Completed"
        Write-AuditEvent -EventType "DataCollection" -Detail "Permissions analysis complete" -AffectedObject $siteUrl
        Complete-AuditSession -Status "Completed"

        # Log throttle stats if any retries occurred
        $throttleStats = Get-ThrottleStats
        if ($throttleStats.TotalRetries -gt 0) {
            Write-ConsoleOutput "⚠️ Throttle protection: $($throttleStats.TotalRetries) retries, $($throttleStats.TotalThrottled) throttle events" -Append
        }

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-RealPermissions-DataDriven"
        Complete-Checkpoint -Status "Failed"
        Write-AuditEvent -EventType "Error" -Detail $_.Exception.Message
        Complete-AuditSession -Status "Failed"
        Write-ConsoleOutput "❌ ERROR: Permissions analysis failed" -Append
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
    }
}

function Get-DemoSites-DataDriven {
    <#
    .SYNOPSIS
    Demo sites using data-driven approach
    #>
    try {
        # Clear and set context
        Clear-SharePointData -DataType "All"
        Set-SharePointOperationContext -OperationType "Sites Analysis (Demo)"

        Write-ConsoleOutput "🎭 Running in Demo Mode - Simulating site collection enumeration..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 500

        # Define demo sites
        $demoSites = @(
            @{Title="Team Collaboration Site"; Url="https://demo.sharepoint.com/sites/teamsite"; Owner="admin@demo.com"; Storage="245"},
            @{Title="Project Alpha Workspace"; Url="https://demo.sharepoint.com/sites/project-alpha"; Owner="project.manager@demo.com"; Storage="1024"},
            @{Title="HR Document Center"; Url="https://demo.sharepoint.com/sites/hr-documents"; Owner="hr.admin@demo.com"; Storage="512"},
            @{Title="Company Intranet Portal"; Url="https://demo.sharepoint.com"; Owner="admin@demo.com"; Storage="2048"},
            @{Title="Sales Team Hub"; Url="https://demo.sharepoint.com/sites/sales"; Owner="sales.manager@demo.com"; Storage="768"}
        )

        Write-ConsoleOutput "🏢 Sites Found: $($demoSites.Count)" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append

        # Add each site to data store
        $siteCounter = 0
        foreach ($site in $demoSites) {
            $siteCounter++
            Add-SharePointSite -SiteData $site

            Write-ConsoleOutput "📁 SITE #$siteCounter`: $($site.Title)" -Append -ForceUpdate
            Write-ConsoleOutput "   🌐 URL: $($site.Url)" -Append
            Write-ConsoleOutput "   👤 Owner: $($site.Owner)" -Append
            Write-ConsoleOutput "   💾 Storage: $($site.Storage) MB" -Append
            Write-ConsoleOutput "   ============================================================" -Append
            Update-UIAndWait -WaitMs 400
        }

        Write-ConsoleOutput "✅ Site enumeration completed successfully!" -Append -ForceUpdate

        # Update Visual Analytics directly from data
        Update-VisualAnalyticsFromData

    }
    catch {
        Write-ConsoleOutput "❌ Demo sites retrieval failed: $($_.Exception.Message)" -Append
    }
}

function Get-DemoPermissions-DataDriven {
    <#
    .SYNOPSIS
    Demo permissions using data-driven approach with realistic storage
    #>
    try {
        # Clear and set context
        Clear-SharePointData -DataType "All"
        Set-SharePointOperationContext -OperationType "Permissions Analysis (Demo)"

        $siteUrl = $script:txtSiteUrl.Text.Trim()
        if ([string]::IsNullOrEmpty($siteUrl)) {
            $siteUrl = "https://demo.sharepoint.com/sites/teamsite"
        }

        # Clear console and show header
        Write-ConsoleOutput "🔐 SHAREPOINT PERMISSIONS ANALYSIS" -ForceUpdate
        Write-ConsoleOutput "=====================================================" -Append -ForceUpdate
        Write-ConsoleOutput "⏱️ Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append -ForceUpdate
        Write-ConsoleOutput "🎯 Target: $siteUrl" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append -ForceUpdate

        Write-ConsoleOutput "🎭 Running in Demo Mode - Simulating permissions analysis..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 500

        # Add demo site with REALISTIC storage value
        $randomStorage = Get-Random -Minimum 200 -Maximum 3000  # Random realistic storage

        Add-SharePointSite -SiteData @{
            Title = "Team Collaboration Site"
            Url = $siteUrl
            Owner = "admin@demo.com"
            Storage = $randomStorage.ToString()  # Use realistic random value
            StorageQuota = "5120"  # 5 GB quota
            HasUniquePermissions = $true
            Template = "STS#3"
            Created = (Get-Date).AddMonths(-6).ToString("yyyy-MM-dd")
            LastModified = (Get-Date).AddDays(-2).ToString("yyyy-MM-dd")
        }

        $storagePercent = [math]::Round(($randomStorage / 5120) * 100, 2)

        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ SITE INFORMATION" -Append -ForceUpdate
        Write-ConsoleOutput "📝 Title: Team Collaboration Site" -Append
        Write-ConsoleOutput "🌐 URL: $siteUrl" -Append
        Write-ConsoleOutput "💾 Storage Used: $randomStorage MB" -Append
        Write-ConsoleOutput "📊 Storage Quota: 5120 MB ($storagePercent% used)" -Append
        Write-ConsoleOutput "🎨 Template: STS#3 (Team Site)" -Append
        Write-ConsoleOutput "📅 Created: $((Get-Date).AddMonths(-6).ToString('yyyy-MM-dd'))" -Append
        Write-ConsoleOutput "📅 Last Modified: $((Get-Date).AddDays(-2).ToString('yyyy-MM-dd'))" -Append
        Write-ConsoleOutput "🔒 Has Unique Permissions: True" -Append
        Write-ConsoleOutput "" -Append

        # Add demo users
        $demoUsers = @(
            @{Name="John Doe"; Email="john.doe@demo.com"; Type="Internal"; Permission="Full Control"; IsSiteAdmin=$true},
            @{Name="Jane Smith"; Email="jane.smith@demo.com"; Type="Internal"; Permission="Edit"; IsSiteAdmin=$false},
            @{Name="Mike Johnson"; Email="mike.johnson@demo.com"; Type="Internal"; Permission="Read"; IsSiteAdmin=$false},
            @{Name="External Partner"; Email="partner@external.com"; Type="External"; Permission="Read"; IsSiteAdmin=$false}
        )

        Write-ConsoleOutput "👤 DETAILED USER PERMISSIONS:" -Append -ForceUpdate
        $userCounter = 0
        foreach ($user in $demoUsers) {
            Add-SharePointUser -UserData $user
            $userCounter++

            $typeIcon = if ($user.Type -eq "External") { "🌐" } else { "🏢" }
            Write-ConsoleOutput "   $typeIcon USER #$userCounter`: $($user.Name)" -Append
            Write-ConsoleOutput "      📧 Email: $($user.Email)" -Append
            Write-ConsoleOutput "      🔒 Permission: $($user.Permission)" -Append
            if ($user.IsSiteAdmin) { Write-ConsoleOutput "      🔑 Site Administrator" -Append }
            Write-ConsoleOutput "   --------------------------------------------------------" -Append
            Update-UIAndWait -WaitMs 200
        }

        Write-ConsoleOutput "" -Append

        # Add demo groups
        $demoGroups = @(
            @{Name="Site Owners"; MemberCount=3; Permission="Full Control"; Description="Site administrators"},
            @{Name="Site Members"; MemberCount=8; Permission="Edit"; Description="Regular contributors"},
            @{Name="Site Visitors"; MemberCount=4; Permission="Read"; Description="Read-only access"}
        )

        Write-ConsoleOutput "👨‍👩‍👧‍👦 SHAREPOINT GROUPS:" -Append -ForceUpdate
        $groupCounter = 0
        foreach ($group in $demoGroups) {
            Add-SharePointGroup -GroupData $group
            $groupCounter++

            Write-ConsoleOutput "   👥 GROUP #$groupCounter`: $($group.Name)" -Append
            Write-ConsoleOutput "      📝 Description: $($group.Description)" -Append
            Write-ConsoleOutput "      👥 Members: $($group.MemberCount)" -Append
            Write-ConsoleOutput "      🔒 Permission: $($group.Permission)" -Append
            Write-ConsoleOutput "   ========================================================" -Append
            Update-UIAndWait -WaitMs 200
        }

        # Add demo role assignments
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "🔑 Analyzing role assignments..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 300

        $demoRoleAssignments = @(
            @{Principal="John Doe"; PrincipalType="User"; Role="Full Control"; Scope="Site"; ScopeUrl=$siteUrl; SiteTitle="Team Collaboration Site"},
            @{Principal="Jane Smith"; PrincipalType="User"; Role="Edit"; Scope="Site"; ScopeUrl=$siteUrl; SiteTitle="Team Collaboration Site"},
            @{Principal="Mike Johnson"; PrincipalType="User"; Role="Read"; Scope="Site"; ScopeUrl=$siteUrl; SiteTitle="Team Collaboration Site"},
            @{Principal="External Partner"; PrincipalType="User"; Role="Read"; Scope="Site"; ScopeUrl=$siteUrl; SiteTitle="Team Collaboration Site"},
            @{Principal="Site Owners"; PrincipalType="SharePoint Group"; Role="Full Control"; Scope="Site"; ScopeUrl=$siteUrl; SiteTitle="Team Collaboration Site"},
            @{Principal="Site Members"; PrincipalType="SharePoint Group"; Role="Edit"; Scope="Site"; ScopeUrl=$siteUrl; SiteTitle="Team Collaboration Site"},
            @{Principal="Site Visitors"; PrincipalType="SharePoint Group"; Role="Read"; Scope="Site"; ScopeUrl=$siteUrl; SiteTitle="Team Collaboration Site"},
            @{Principal="HR Team"; PrincipalType="Security Group"; Role="Contribute"; Scope="Library"; ScopeUrl="/sites/teamsite/Shared Documents"; SiteTitle="Team Collaboration Site"},
            @{Principal="Finance Group"; PrincipalType="Security Group"; Role="Read"; Scope="List"; ScopeUrl="/sites/teamsite/Lists/Budget"; SiteTitle="Team Collaboration Site"},
            @{Principal="IT Admins"; PrincipalType="Security Group"; Role="Full Control"; Scope="Site"; ScopeUrl=$siteUrl; SiteTitle="Team Collaboration Site"},
            @{Principal="Marketing Team"; PrincipalType="Security Group"; Role="Edit"; Scope="Library"; ScopeUrl="/sites/teamsite/Marketing Assets"; SiteTitle="Team Collaboration Site"},
            @{Principal="Contractors"; PrincipalType="SharePoint Group"; Role="View Only"; Scope="Library"; ScopeUrl="/sites/teamsite/Project Files"; SiteTitle="Team Collaboration Site"}
        )

        foreach ($ra in $demoRoleAssignments) {
            Add-SharePointRoleAssignment -RoleData $ra
        }

        Write-ConsoleOutput "✅ Found $($demoRoleAssignments.Count) role assignments" -Append -ForceUpdate

        # Add demo inheritance items
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "🔗 Checking permission inheritance..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 300

        $demoInheritanceItems = @(
            @{Title="Team Collaboration Site"; Type="Site"; Url=$siteUrl; HasUniquePermissions=$true; ParentUrl="N/A"; RoleAssignmentCount=7; SiteTitle="Team Collaboration Site"},
            @{Title="Shared Documents"; Type="Document Library"; Url="/sites/teamsite/Shared Documents"; HasUniquePermissions=$true; ParentUrl=$siteUrl; RoleAssignmentCount=4; SiteTitle="Team Collaboration Site"},
            @{Title="Site Pages"; Type="Document Library"; Url="/sites/teamsite/SitePages"; HasUniquePermissions=$false; ParentUrl=$siteUrl; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"},
            @{Title="Project Files"; Type="Document Library"; Url="/sites/teamsite/Project Files"; HasUniquePermissions=$true; ParentUrl=$siteUrl; RoleAssignmentCount=5; SiteTitle="Team Collaboration Site"},
            @{Title="Marketing Assets"; Type="Document Library"; Url="/sites/teamsite/Marketing Assets"; HasUniquePermissions=$true; ParentUrl=$siteUrl; RoleAssignmentCount=3; SiteTitle="Team Collaboration Site"},
            @{Title="Budget Tracker"; Type="List"; Url="/sites/teamsite/Lists/Budget"; HasUniquePermissions=$true; ParentUrl=$siteUrl; RoleAssignmentCount=2; SiteTitle="Team Collaboration Site"},
            @{Title="Team Calendar"; Type="List"; Url="/sites/teamsite/Lists/Calendar"; HasUniquePermissions=$false; ParentUrl=$siteUrl; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"},
            @{Title="Announcements"; Type="List"; Url="/sites/teamsite/Lists/Announcements"; HasUniquePermissions=$false; ParentUrl=$siteUrl; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"},
            @{Title="Style Library"; Type="Document Library"; Url="/sites/teamsite/Style Library"; HasUniquePermissions=$false; ParentUrl=$siteUrl; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"},
            @{Title="Form Templates"; Type="Document Library"; Url="/sites/teamsite/FormServerTemplates"; HasUniquePermissions=$false; ParentUrl=$siteUrl; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"}
        )

        foreach ($item in $demoInheritanceItems) {
            Add-SharePointInheritanceItem -InheritanceData $item
        }

        $brokenCount = ($demoInheritanceItems | Where-Object { $_["HasUniquePermissions"] -eq $true }).Count
        Write-ConsoleOutput "✅ Scanned $($demoInheritanceItems.Count) items - $brokenCount with broken inheritance" -Append -ForceUpdate

        # Add demo sharing links
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "🔗 Auditing sharing links..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 300

        $demoSharingLinks = @(
            @{GroupName="SharingLinks.abc123.OrganizationView.def456"; LinkType="Company-wide"; AccessLevel="View"; MemberCount=0; SiteTitle="Team Collaboration Site"; CreatedDate="2025-01-15"},
            @{GroupName="SharingLinks.abc124.OrganizationEdit.ghi789"; LinkType="Company-wide"; AccessLevel="Edit"; MemberCount=0; SiteTitle="Team Collaboration Site"; CreatedDate="2025-02-01"},
            @{GroupName="SharingLinks.abc125.AnonymousView.jkl012"; LinkType="Anonymous"; AccessLevel="View"; MemberCount=3; SiteTitle="Team Collaboration Site"; CreatedDate="2025-01-20"},
            @{GroupName="SharingLinks.abc126.Flexible.mno345"; LinkType="Specific People"; AccessLevel="Edit"; MemberCount=5; SiteTitle="Team Collaboration Site"; CreatedDate="2025-02-10"},
            @{GroupName="SharingLinks.abc127.Flexible.pqr678"; LinkType="Specific People"; AccessLevel="View"; MemberCount=2; SiteTitle="Team Collaboration Site"; CreatedDate="2025-01-25"},
            @{GroupName="SharingLinks.abc128.AnonymousEdit.stu901"; LinkType="Anonymous"; AccessLevel="Edit"; MemberCount=1; SiteTitle="Team Collaboration Site"; CreatedDate="2025-03-01"},
            @{GroupName="SharingLinks.abc129.Flexible.vwx234"; LinkType="Specific People"; AccessLevel="Edit"; MemberCount=8; SiteTitle="Team Collaboration Site"; CreatedDate="2025-02-15"}
        )

        foreach ($link in $demoSharingLinks) {
            Add-SharePointSharingLink -LinkData $link
        }

        Write-ConsoleOutput "✅ Found $($demoSharingLinks.Count) sharing links" -Append -ForceUpdate

        # Display summary
        $metrics = Get-SharePointData -DataType "Metrics"
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "📊 PERMISSION SUMMARY:" -Append -ForceUpdate
        Write-ConsoleOutput "   👤 Total Users: $($metrics.TotalUsers)" -Append
        Write-ConsoleOutput "   👨‍👩‍👧‍👦 Total Groups: $($metrics.TotalGroups)" -Append
        Write-ConsoleOutput "   🌐 External Users: $($metrics.ExternalUsers)" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "🔑 CORE ANALYSIS:" -Append -ForceUpdate
        Write-ConsoleOutput "   • Role Assignments: $($metrics.TotalRoleAssignments)" -Append
        Write-ConsoleOutput "   • Inheritance Breaks: $($metrics.InheritanceBreaks)" -Append
        Write-ConsoleOutput "   • Sharing Links: $($metrics.TotalSharingLinks)" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ Permissions analysis completed successfully!" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "💡 TIP: Click metric tiles on Visual Analytics tab for deep dive analysis" -Append -ForceUpdate
        Write-ConsoleOutput "   New: Role Assignments, Inheritance Tree, and Sharing Links deep dives!" -Append

        # Update Visual Analytics directly from data
        Update-VisualAnalyticsFromData

    }
    catch {
        Write-ConsoleOutput "❌ Demo permissions analysis failed: $($_.Exception.Message)" -Append
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-DemoPermissions-DataDriven"
    }
}

function Generate-SharePointReport {
    <#
    .SYNOPSIS
    Generates comprehensive SharePoint permissions report with real-time progress
    #>
    try {
        Write-ActivityLog "Starting SharePoint report generation" -Level "Information"

        # Clear previous results and show starting message
        Write-ConsoleOutput "📈 SHAREPOINT PERMISSIONS REPORT GENERATION" -ForceUpdate
        Write-ConsoleOutput "=======================================================" -Append -ForceUpdate
        Write-ConsoleOutput "⏱️ Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append -ForceUpdate

        if (Get-AppSetting -SettingName "DemoMode") {
            Generate-DemoReport
        } else {
            Generate-RealReport
        }

        Write-ActivityLog "SharePoint report generation completed" -Level "Information"
    }
    catch {
        Write-ConsoleOutput "❌ ERROR: Report generation failed!" -Append
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
        Write-ErrorLog -Message $_.Exception.Message -Location "Generate-SharePointReport"
    }
}

function Generate-DemoReport {
    <#
    .SYNOPSIS
    Generates demo report with realistic progress simulation
    #>
    try {
        # Initialize progress tracking
        $script:ReportStartTime = Get-Date

        Write-ConsoleOutput "🎭 Running in Demo Mode - Simulating comprehensive report generation..." -Append -ForceUpdate
        Write-ConsoleOutput "" -Append

        # Step-by-step progress
        Show-ReportProgress "Initializing Report Environment" "Creating output directories and validating settings" 600
        Show-ReportProgress "Gathering Site Collections" "Enumerating all SharePoint sites in tenant" 1200
        Show-ReportProgress "Analyzing User Permissions" "Processing individual user access rights" 1500
        Show-ReportProgress "Processing Group Memberships" "Mapping SharePoint groups and nested permissions" 900
        Show-ReportProgress "Detecting Security Issues" "Scanning for orphaned permissions and risks" 700
        Show-ReportProgress "Generating CSV Files" "Creating detailed permission matrices" 800
        Show-ReportProgress "Creating Excel Workbook" "Formatting data with charts and pivot tables" 600
        Show-ReportProgress "Finalizing Reports" "Adding executive summary and recommendations" 400

        # Generate results
        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $totalTime = [math]::Round(((Get-Date) - $script:ReportStartTime).TotalSeconds, 1)

        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ REPORT GENERATION COMPLETED!" -Append -ForceUpdate
        Write-ConsoleOutput "========================================" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "📁 GENERATED FILES:" -Append -ForceUpdate
        Write-ConsoleOutput "   📊 SharePoint_Permissions_Report_$timestamp.csv" -Append
        Write-ConsoleOutput "   📈 SharePoint_Analysis_Dashboard_$timestamp.xlsx" -Append
        Write-ConsoleOutput "   📋 Executive_Summary_$timestamp.pdf" -Append
        Write-ConsoleOutput "   📝 Detailed_Audit_Log_$timestamp.txt" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "📊 REPORT STATISTICS:" -Append -ForceUpdate
        Write-ConsoleOutput "   ⏱️ Processing Time: $totalTime seconds" -Append
        Write-ConsoleOutput "   🏢 Sites Analyzed: 5" -Append
        Write-ConsoleOutput "   👥 Users Processed: 47" -Append
        Write-ConsoleOutput "   👨‍👩‍👧‍👦 Groups Analyzed: 23" -Append
        Write-ConsoleOutput "   📋 Permission Records: 1,247" -Append
        Write-ConsoleOutput "   🔍 Security Findings: 8" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "🎯 KEY FINDINGS:" -Append -ForceUpdate
        Write-ConsoleOutput "   ⚠️ 8 External users detected (review required)" -Append
        Write-ConsoleOutput "   🔴 3 Sites with excessive storage usage" -Append
        Write-ConsoleOutput "   ⚠️ 2 Orphaned permission groups found" -Append
        Write-ConsoleOutput "   ✅ All sites comply with base governance policies" -Append
        Write-ConsoleOutput "   🔍 5 Permission inheritance breaks detected" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "📋 NEXT STEPS:" -Append -ForceUpdate
        Write-ConsoleOutput "   1. Review generated CSV files for detailed analysis" -Append
        Write-ConsoleOutput "   2. Open Excel dashboard for visual insights" -Append
        Write-ConsoleOutput "   3. Address identified security recommendations" -Append
        Write-ConsoleOutput "   4. Schedule regular automated permission audits" -Append
        Write-ConsoleOutput "   5. Implement governance policies for future prevention" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "💡 TIP: Check the Visual Analytics tab for an interactive overview!" -Append -ForceUpdate

        # Update Visual Analytics using the data-driven approach
        Update-VisualAnalyticsFromData
    }
    catch {
        throw "Demo report generation failed: $($_.Exception.Message)"
    }
}

function Generate-RealReport {
    <#
    .SYNOPSIS
    Generates real SharePoint report using actual data
    #>
    try {
        if (-not $script:SPOConnected) {
            throw "Not connected to SharePoint. Please connect first."
        }

        Write-ConsoleOutput "🔄 Preparing real SharePoint report generation..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 1000

        Write-ConsoleOutput "📊 Gathering comprehensive tenant data..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 1500

        # Initialize progress tracking
        $script:ReportStartTime = Get-Date

        # Step 1: Gather Sites Data
        Show-ReportProgress "Gathering Site Collections" "Enumerating all accessible SharePoint sites" 1000

        $allSites = @()
        try {
            # Try to get tenant sites first
            try {
                $allSites = Get-PnPTenantSite -ErrorAction Stop
                Write-ConsoleOutput "   ✅ Retrieved $($allSites.Count) sites from tenant level" -Append
            }
            catch {
                # Fallback to current site and subsites
                $currentWeb = Get-PnPWeb -ErrorAction Stop
                $subsites = Get-PnPSubWeb -Recurse -ErrorAction SilentlyContinue
                $allSites = @($currentWeb)
                if ($subsites) { $allSites += $subsites }
                Write-ConsoleOutput "   ⚠️ Limited to accessible sites: $($allSites.Count) found" -Append
            }
        }
        catch {
            Write-ConsoleOutput "   ❌ Failed to gather sites data: $($_.Exception.Message)" -Append
            throw "Cannot generate report without site data"
        }

        # Step 2: Analyze Permissions for Each Site
        Show-ReportProgress "Analyzing Site Permissions" "Processing permissions for each site" 2000

        $allPermissions = @()
        $siteCounter = 0
        foreach ($site in $allSites | Select-Object -First 10) { # Limit for performance
            $siteCounter++
            try {
                Write-ConsoleOutput "   🔍 Analyzing site $siteCounter/$($allSites.Count): $($site.Title)" -Append

                # Connect to each site
                Connect-PnPOnline -Url $site.Url -ClientId (Get-AppSetting -SettingName "SharePoint.ClientId") -Interactive -ErrorAction SilentlyContinue

                # Get users and groups for this site
                $users = Get-PnPUser -ErrorAction SilentlyContinue
                $groups = Get-PnPGroup -ErrorAction SilentlyContinue

                # Process users
                foreach ($user in $users | Where-Object { $_.PrincipalType -eq "User" }) {
                    $allPermissions += [PSCustomObject]@{
                        SiteTitle = $site.Title
                        SiteUrl = $site.Url
                        PrincipalName = $user.Title
                        PrincipalEmail = $user.Email
                        PrincipalType = "User"
                        LoginName = $user.LoginName
                        IsSiteAdmin = $user.IsSiteAdmin
                        IsExternal = $user.IsShareByEmailGuestUser -or $user.IsEmailAuthenticationGuestUser
                        LastModified = Get-Date
                    }
                }

                # Process groups
                foreach ($group in $groups) {
                    $allPermissions += [PSCustomObject]@{
                        SiteTitle = $site.Title
                        SiteUrl = $site.Url
                        PrincipalName = $group.Title
                        PrincipalEmail = ""
                        PrincipalType = "Group"
                        LoginName = $group.LoginName
                        IsSiteAdmin = $false
                        IsExternal = $false
                        LastModified = Get-Date
                    }
                }

                Update-UIAndWait -WaitMs 200
            }
            catch {
                Write-ConsoleOutput "   ⚠️ Skipped site due to access issues: $($site.Title)" -Append
            }
        }

        # Step 3: Generate CSV Report
        Show-ReportProgress "Generating CSV Export" "Creating detailed permission matrices" 800

        $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $csvPath = "SharePoint_Permissions_Report_$timestamp.csv"

        try {
            $allPermissions | Export-Csv -Path $csvPath -NoTypeInformation -ErrorAction Stop
            Write-ConsoleOutput "   ✅ CSV report saved: $csvPath" -Append
        }
        catch {
            Write-ConsoleOutput "   ❌ Failed to save CSV: $($_.Exception.Message)" -Append
        }

        # Step 4: Generate Summary Statistics
        Show-ReportProgress "Generating Summary Statistics" "Calculating security metrics and insights" 600

        $totalUsers = ($allPermissions | Where-Object { $_.PrincipalType -eq "User" }).Count
        $totalGroups = ($allPermissions | Where-Object { $_.PrincipalType -eq "Group" }).Count
        $externalUsers = ($allPermissions | Where-Object { $_.IsExternal -eq $true }).Count
        $siteAdmins = ($allPermissions | Where-Object { $_.IsSiteAdmin -eq $true }).Count

        # Step 5: Create Executive Summary
        Show-ReportProgress "Creating Executive Summary" "Compiling findings and recommendations" 400

        $totalTime = [math]::Round(((Get-Date) - $script:ReportStartTime).TotalSeconds, 1)

        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ REAL SHAREPOINT REPORT COMPLETED!" -Append -ForceUpdate
        Write-ConsoleOutput "========================================" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "📁 GENERATED FILES:" -Append -ForceUpdate
        Write-ConsoleOutput "   📊 $csvPath" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "📊 REPORT STATISTICS:" -Append -ForceUpdate
        Write-ConsoleOutput "   ⏱️ Processing Time: $totalTime seconds" -Append
        Write-ConsoleOutput "   🏢 Sites Analyzed: $siteCounter" -Append
        Write-ConsoleOutput "   👥 Total Users: $totalUsers" -Append
        Write-ConsoleOutput "   👨‍👩‍👧‍👦 Total Groups: $totalGroups" -Append
        Write-ConsoleOutput "   📋 Permission Records: $($allPermissions.Count)" -Append
        Write-ConsoleOutput "   🌐 External Users: $externalUsers" -Append
        Write-ConsoleOutput "   🔑 Site Administrators: $siteAdmins" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "🎯 KEY FINDINGS:" -Append -ForceUpdate
        if ($externalUsers -gt 0) {
            Write-ConsoleOutput "   ⚠️ $externalUsers External users detected (review required)" -Append
        }
        if ($siteAdmins -gt 5) {
            Write-ConsoleOutput "   ⚠️ $siteAdmins Site administrators found (review recommended)" -Append
        }
        Write-ConsoleOutput "   ✅ Permissions data successfully exported for detailed analysis" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "📋 NEXT STEPS:" -Append -ForceUpdate
        Write-ConsoleOutput "   1. Open the generated CSV file for detailed analysis" -Append
        Write-ConsoleOutput "   2. Review external user access and justifications" -Append
        Write-ConsoleOutput "   3. Audit site administrator assignments" -Append
        Write-ConsoleOutput "   4. Consider implementing automated monitoring" -Append
        Write-ConsoleOutput "   5. Schedule regular permission reviews" -Append
        Write-ConsoleOutput "" -Append

        Write-ConsoleOutput "💡 CSV file contains complete permission data for further analysis in Excel or Power BI!" -Append -ForceUpdate

        # Update Visual Analytics using the data-driven approach
        Update-VisualAnalyticsFromData

    }
    catch {
        Write-ConsoleOutput "❌ Real SharePoint report generation failed!" -Append
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "📝 To implement full report generation:" -Append
        Write-ConsoleOutput "   1. Ensure you have SharePoint Administrator permissions" -Append
        Write-ConsoleOutput "   2. Verify app registration has appropriate API permissions" -Append
        Write-ConsoleOutput "   3. Test with Demo Mode to see expected functionality" -Append
        Write-ConsoleOutput "   4. Consider using PowerShell execution policies if file creation fails" -Append
        throw "Real report generation failed: $($_.Exception.Message)"
    }
}

function Show-ReportProgress {
    <#
    .SYNOPSIS
    Shows individual report generation progress steps
    #>
    param(
        [string]$StepName,
        [string]$Details,
        [int]$DelayMs = 500
    )

    try {
        $timeStamp = Get-Date -Format "HH:mm:ss"

        Write-ConsoleOutput "[$timeStamp] 🔄 $StepName" -Append -ForceUpdate
        if ($Details) {
            Write-ConsoleOutput "           └─ $Details" -Append
        }

        Update-UIAndWait -WaitMs $DelayMs

        Write-ConsoleOutput "[$timeStamp] ✅ $StepName - Complete" -Append
        Write-ConsoleOutput "" -Append
    }
    catch {
        Write-ActivityLog "Error showing report progress: $($_.Exception.Message)" -Level "Warning"
    }
}

# [MARKED FOR DELETION] Duplicate of Update-UIAndWait in MainWindow.ps1
# The MainWindow.ps1 version is more complete (uses WPF Dispatcher.Invoke in addition to
# DoEvents, defaults $WaitMs=0, conditional sleep). This WinForms-only variant was introduced
# in the same commit ("Mainwindow logic split", 2025-08-11) and is already overridden at
# runtime since MainWindow.ps1 loads after OperationsTab.ps1.
# function Update-UIAndWait {
#     param([int]$WaitMs = 500)
#
#     try {
#         # Force UI update
#         [System.Windows.Forms.Application]::DoEvents()
#
#         # Wait for specified time
#         Start-Sleep -Milliseconds $WaitMs
#     }
#     catch {
#         # Ignore errors in UI updates
#         Start-Sleep -Milliseconds $WaitMs
#     }
# }

# [MARKED FOR DELETION] Duplicate of Write-ConsoleOutput in MainWindow.ps1
# The MainWindow.ps1 version uses correct WPF APIs (.ScrollToEnd(), Dispatcher-based refresh,
# $NewLine parameter). This variant uses WinForms APIs (.SelectionStart/.ScrollToCaret()) that
# are incompatible with the WPF TextBox. Introduced in the same commit ("Mainwindow logic split",
# 2025-08-11) and already overridden at runtime since MainWindow.ps1 loads after OperationsTab.ps1.
# function Write-ConsoleOutput {
#     param(
#         [string]$Message,
#         [switch]$Append,
#         [switch]$ForceUpdate
#     )
#
#     try {
#         if ($Append) {
#             $script:txtOperationsResults.Text += "`n$Message"
#         } else {
#             $script:txtOperationsResults.Text = $Message
#         }
#
#         if ($ForceUpdate) {
#             # Scroll to bottom
#             $script:txtOperationsResults.SelectionStart = $script:txtOperationsResults.Text.Length
#             $script:txtOperationsResults.ScrollToCaret()
#
#             # Force UI refresh
#             [System.Windows.Forms.Application]::DoEvents()
#         }
#     }
#     catch {
#         # Fallback to Write-Host if UI is not available
#         Write-Host $Message
#     }
# }
