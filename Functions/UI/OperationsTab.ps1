# ============================================
# OperationsTab.ps1 - Real SharePoint Operations
# ============================================
# Contains functions for live SharePoint data collection.
# These run inside background runspaces (Start-BackgroundOperation)
# where Write-ConsoleOutput writes to the shared operation log
# and Update-UIAndWait is a no-op.

function Get-RealSites-DataDriven {
    <#
    .SYNOPSIS
    Retrieves real SharePoint sites with proper storage data
    #>
    try {
        # Clear previous sites data and set context
        Clear-SharePointData -DataType "Sites"
        Set-SharePointOperationContext -OperationType "Sites Analysis"
        Start-Checkpoint -OperationType "SitesAnalysis" -Scope "Tenant"
        Reset-ThrottleStats

        # Start audit session
        Start-AuditSession -OperationType "SitesAnalysis" -ScanScope "Tenant"

        Write-ConsoleOutput "SHAREPOINT SITES ANALYSIS"
        Write-ConsoleOutput "====================================================="
        Write-ConsoleOutput "Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        Write-ConsoleOutput ""

        Write-ConsoleOutput "Using modern PnP PowerShell for site enumeration..."
        Write-ConsoleOutput "Attempting tenant-level site enumeration..."

        $sites = @()
        $requiresAdmin = $false

        # Get tenant URL and client ID for reconnection
        $tenantUrl = Get-AppSetting -SettingName "SharePoint.TenantUrl"
        $adminUrl = $tenantUrl -replace "\.sharepoint\.com", "-admin.sharepoint.com"
        $clientId = Get-AppSetting -SettingName "SharePoint.ClientId"

        # Try to get sites with full details using the existing connection
        try {
            Write-ConsoleOutput "Scanning tenant for site collections..."

            # First try Get-PnPTenantSite on the existing connection (may need admin)
            try {
                Write-ConsoleOutput "Attempting to enumerate tenant sites..."
                $sites = Invoke-WithThrottleProtection -OperationName "Get-PnPTenantSite" -ScriptBlock {
                    Get-PnPTenantSite -ErrorAction Stop
                }
                Write-ActivityLog "Retrieved $($sites.Count) sites from tenant"
                Write-ConsoleOutput "Successfully retrieved $($sites.Count) sites"
            }
            catch {
                Write-ActivityLog "Standard enumeration failed, trying admin connection: $($_.Exception.Message)"
                Write-ConsoleOutput "Standard access limited, trying admin center..."

                # Fallback to admin connection
                $token = $null
                try { $token = Get-PnPAccessToken -ErrorAction SilentlyContinue } catch { }

                if ($token) {
                    Connect-PnPOnline -Url $adminUrl -AccessToken $token -ErrorAction Stop
                } else {
                    Connect-PnPOnline -Url $adminUrl -ClientId $clientId -Interactive -ErrorAction Stop
                }
                $sites = Invoke-WithThrottleProtection -OperationName "Get-PnPTenantSite (admin)" -ScriptBlock {
                    Get-PnPTenantSite -Detailed -ErrorAction Stop
                }
                $requiresAdmin = $true
                Write-ActivityLog "Retrieved $($sites.Count) sites from admin center"
                Write-ConsoleOutput "Successfully retrieved $($sites.Count) sites with storage details"
            }
        }
        catch {
            Write-ActivityLog "Tenant-level enumeration failed: $($_.Exception.Message)"
            Write-ConsoleOutput "Tenant-level access limited, using fallback method..."

            # Final fallback - get current site only
            try {
                $currentWeb = Get-PnPWeb -ErrorAction Stop
                $currentSite = Get-PnPSite -Includes Usage -ErrorAction Stop

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
                Write-ConsoleOutput "Found current site only"
            }
            catch {
                throw "Unable to retrieve any sites. SharePoint Administrator permissions may be required."
            }
        }

        Write-ConsoleOutput ""
        Write-ConsoleOutput "SITES DISCOVERY RESULTS"
        Write-ConsoleOutput "Sites Found: $($sites.Count)"
        Write-ConsoleOutput ""

        # Process and store each site with proper storage data
        $siteCounter = 0
        foreach ($site in $sites | Select-Object -First 25) {
            $siteCounter++

            # Extract storage value properly based on the object type
            $storageValue = 0

            if ($null -ne $site.StorageUsageCurrent) {
                $storageValue = [int]$site.StorageUsageCurrent
            }
            elseif ($site.Usage -and $site.Usage.Storage) {
                $storageValue = [math]::Round($site.Usage.Storage / 1MB, 0)
            }
            elseif ($site.Url) {
                try {
                    $token = $null
                    try { $token = Get-PnPAccessToken -ErrorAction SilentlyContinue } catch { }
                    if ($token) {
                        Connect-PnPOnline -Url $site.Url -AccessToken $token -ErrorAction SilentlyContinue
                    }
                    $siteDetail = Get-PnPSite -Includes Usage -ErrorAction SilentlyContinue
                    if ($siteDetail -and $siteDetail.Usage -and $siteDetail.Usage.Storage) {
                        $storageValue = [math]::Round($siteDetail.Usage.Storage / 1MB, 0)
                    }
                }
                catch {
                    Write-ActivityLog "Could not get storage for site: $($site.Url)" -Level "Warning"
                }
            }

            $siteData = @{
                Title = if ($site.Title) { $site.Title } else { "Site $siteCounter" }
                Url = if ($site.Url) { $site.Url } else { "N/A" }
                Owner = if ($site.Owner) { $site.Owner } elseif ($site.SiteOwnerEmail) { $site.SiteOwnerEmail } else { "N/A" }
                Storage = $storageValue.ToString()
                StorageQuota = if ($site.StorageQuota) { $site.StorageQuota } else { 0 }
                Template = if ($site.Template) { $site.Template } else { "N/A" }
                LastModified = if ($site.LastContentModifiedDate) {
                    $site.LastContentModifiedDate.ToString("yyyy-MM-dd")
                } else { "N/A" }
                IsHubSite = if ($site.IsHubSite) { $site.IsHubSite } else { $false }
                UserCount = 0
                GroupCount = 0
            }

            Add-SharePointSite -SiteData $siteData

            Write-ConsoleOutput "SITE #${siteCounter}: $($siteData.Title)"
            Write-ConsoleOutput "   URL: $($siteData.Url)"
            Write-ConsoleOutput "   Owner: $($siteData.Owner)"
            Write-ConsoleOutput "   Storage: $storageValue MB"
            if ($siteData.StorageQuota -gt 0) {
                $quotaMB = [math]::Round($siteData.StorageQuota / 1MB, 0)
                $percentUsed = if ($quotaMB -gt 0) {
                    [math]::Round(($storageValue / $quotaMB) * 100, 1)
                } else { 0 }
                Write-ConsoleOutput "   Storage Quota: $quotaMB MB ($percentUsed% used)"
            }
            if ($siteData.Template -ne "N/A") { Write-ConsoleOutput "   Template: $($siteData.Template)" }
            if ($siteData.LastModified -ne "N/A") { Write-ConsoleOutput "   Last Modified: $($siteData.LastModified)" }
        }

        if ($sites.Count -gt 25) {
            Write-ConsoleOutput "... and $($sites.Count - 25) more sites"
            Write-ConsoleOutput ""
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
            Write-ConsoleOutput ""
            Write-ConsoleOutput "STORAGE SUMMARY:"
            Write-ConsoleOutput "   Total Storage Used: $totalStorage MB ($totalGB GB)"
            Write-ConsoleOutput "   Average per Site: $avgMB MB"
            Write-ConsoleOutput "   Sites with Storage Data: $sitesWithStorage/$($sites.Count)"
        }

        if (-not $requiresAdmin -and $sitesWithStorage -eq 0) {
            Write-ConsoleOutput ""
            Write-ConsoleOutput "Note: Storage data requires SharePoint Administrator permissions"
        }

        Write-ConsoleOutput ""
        Write-ConsoleOutput "Site enumeration completed successfully!"

        Write-ActivityLog "Sites operation completed with $($sites.Count) sites" -Level "Information"
        Complete-Checkpoint -Status "Completed"
        Write-AuditEvent -EventType "DataCollection" -Detail "Sites analysis complete: $($sites.Count) sites"
        Complete-AuditSession -Status "Completed"

        # Log throttle stats if any retries occurred
        $throttleStats = Get-ThrottleStats
        if ($throttleStats.TotalRetries -gt 0) {
            Write-ConsoleOutput "Throttle protection: $($throttleStats.TotalRetries) retries, $($throttleStats.TotalThrottled) throttle events"
        }
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-RealSites-DataDriven"
        Complete-Checkpoint -Status "Failed"
        Write-AuditEvent -EventType "Error" -Detail $_.Exception.Message
        Complete-AuditSession -Status "Failed"
        Write-ConsoleOutput ""
        Write-ConsoleOutput "ERROR: Site enumeration failed"
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)"
    }
}

function Get-RealPermissions-DataDriven {
    <#
    .SYNOPSIS
    Analyzes real SharePoint permissions and stores in data manager
    .PARAMETER SiteUrl
    The SharePoint site URL to analyze permissions for
    #>
    param(
        [Parameter(Mandatory = $false)]
        [string]$SiteUrl
    )

    try {
        if ([string]::IsNullOrEmpty($SiteUrl)) {
            Write-ConsoleOutput "No site URL provided. Please specify a site URL to analyze permissions."
            return
        }

        # Clear previous data and set context
        Clear-SharePointData -DataType "All"
        Set-SharePointOperationContext -OperationType "Permissions Analysis"
        Start-Checkpoint -OperationType "PermissionsAnalysis" -Scope $SiteUrl
        Reset-ThrottleStats

        # Start audit session
        Start-AuditSession -OperationType "PermissionsAnalysis" -ScanScope $SiteUrl

        Write-ConsoleOutput "SHAREPOINT PERMISSIONS ANALYSIS"
        Write-ConsoleOutput "====================================================="
        Write-ConsoleOutput "Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        Write-ConsoleOutput "Target: $SiteUrl"
        Write-ConsoleOutput ""

        Write-ConsoleOutput "Analyzing permissions for: $SiteUrl..."

        # Check if the background runspace already connected to the target site.
        # If not (or connected to a different URL), reconnect.
        $currentUrl = $null
        try {
            $currentUrl = (Get-PnPConnection -ErrorAction SilentlyContinue).Url
        } catch { }

        if ($currentUrl -and $currentUrl.TrimEnd('/') -eq $SiteUrl.TrimEnd('/')) {
            Write-ConsoleOutput "Using existing connection to site..."
        } else {
            Write-ConsoleOutput "Connecting to the specified site..."
            $token = $null
            try { $token = Get-PnPAccessToken -ErrorAction SilentlyContinue } catch { }
            if ($token) {
                Connect-PnPOnline -Url $SiteUrl -AccessToken $token -ErrorAction Stop
            } else {
                $cid = Get-AppSetting -SettingName "SharePoint.ClientId"
                Connect-PnPOnline -Url $SiteUrl -ClientId $cid -Interactive
            }
        }
        Write-ConsoleOutput "Connected successfully!"

        # Get site information INCLUDING STORAGE
        Write-ConsoleOutput "Retrieving site information and storage data..."

        $web = Get-PnPWeb -ErrorAction Stop

        # Get REAL storage data
        $storageValue = 0
        $storageQuota = 0
        $storagePercentUsed = 0

        try {
            $site = Get-PnPSite -Includes Usage -ErrorAction Stop

            if ($site.Usage) {
                $storageValue = [math]::Round($site.Usage.Storage / 1MB, 0)
                $storageQuota = [math]::Round($site.Usage.StorageQuotaInMB, 0)

                if ($site.Usage.StoragePercentageUsed) {
                    $storagePercentUsed = [math]::Round($site.Usage.StoragePercentageUsed, 2)
                }

                Write-ActivityLog "Retrieved storage data: $storageValue MB used of $storageQuota MB quota ($storagePercentUsed%)" -Level "Information"
            }
        }
        catch {
            Write-ActivityLog "Failed to get storage data: $($_.Exception.Message)" -Level "Warning"

            # Try alternative method using tenant admin if available
            try {
                $tenantUrl = Get-AppSetting -SettingName "SharePoint.TenantUrl"
                $adminUrl = $tenantUrl -replace "\.sharepoint\.com", "-admin.sharepoint.com"

                Write-ConsoleOutput "Attempting to get storage via admin connection..."

                $adminToken = $null
                try { $adminToken = Get-PnPAccessToken -ErrorAction SilentlyContinue } catch { }
                if ($adminToken) {
                    Connect-PnPOnline -Url $adminUrl -AccessToken $adminToken -ErrorAction Stop
                } else {
                    $cid = Get-AppSetting -SettingName "SharePoint.ClientId"
                    Connect-PnPOnline -Url $adminUrl -ClientId $cid -Interactive -ErrorAction Stop
                }
                $tenantSites = Get-PnPTenantSite -Url $SiteUrl -Detailed -ErrorAction Stop

                if ($tenantSites -and $tenantSites.StorageUsageCurrent) {
                    $storageValue = [int]$tenantSites.StorageUsageCurrent
                    $storageQuota = if ($tenantSites.StorageQuota) { [int]$tenantSites.StorageQuota } else { 0 }
                    Write-ActivityLog "Retrieved storage from admin: $storageValue MB" -Level "Information"
                }

                # Reconnect to the original site
                $siteToken = $null
                try { $siteToken = Get-PnPAccessToken -ErrorAction SilentlyContinue } catch { }
                if ($siteToken) {
                    Connect-PnPOnline -Url $SiteUrl -AccessToken $siteToken -ErrorAction Stop
                } else {
                    $cid = Get-AppSetting -SettingName "SharePoint.ClientId"
                    Connect-PnPOnline -Url $SiteUrl -ClientId $cid -Interactive
                }
            }
            catch {
                Write-ActivityLog "Admin storage retrieval also failed: $($_.Exception.Message)" -Level "Warning"
            }
        }

        # Add the analyzed site to data store with REAL storage
        $siteData = @{
            Title = if ($web.Title) { $web.Title } else { "Analyzed Site" }
            Url = if ($web.Url) { $web.Url } else { $SiteUrl }
            Owner = "Current User"
            Storage = $storageValue.ToString()
            StorageQuota = $storageQuota.ToString()
            Template = if ($web.WebTemplate) { $web.WebTemplate } else { "N/A" }
            HasUniquePermissions = $web.HasUniqueRoleAssignments
            Created = if ($web.Created) { $web.Created.ToString("yyyy-MM-dd") } else { "N/A" }
            LastModified = if ($web.LastItemModifiedDate) { $web.LastItemModifiedDate.ToString("yyyy-MM-dd") } else { "N/A" }
        }
        Add-SharePointSite -SiteData $siteData

        Write-ConsoleOutput ""
        Write-ConsoleOutput "SITE INFORMATION"
        Write-ConsoleOutput "Title: $($web.Title)"
        Write-ConsoleOutput "URL: $($web.Url)"
        Write-ConsoleOutput "Storage Used: $storageValue MB"
        if ($storageQuota -gt 0) {
            Write-ConsoleOutput "Storage Quota: $storageQuota MB ($storagePercentUsed% used)"
        }
        Write-ConsoleOutput "Template: $($web.WebTemplate)"
        Write-ConsoleOutput "Created: $($siteData.Created)"
        Write-ConsoleOutput "Last Modified: $($siteData.LastModified)"
        Write-ConsoleOutput "Has Unique Permissions: $($web.HasUniqueRoleAssignments)"
        Write-ConsoleOutput ""

        # Get and store users
        Write-ConsoleOutput "Retrieving users..."
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
            }

            Write-ConsoleOutput "Retrieved $($regularUsers.Count) users"
            Write-ConsoleOutput ""
        }
        catch {
            Write-ConsoleOutput "Limited access to user information: $($_.Exception.Message)"
            Write-ConsoleOutput ""
        }

        # Get and store groups
        Write-ConsoleOutput "Retrieving groups..."
        Update-Checkpoint -Phase "Groups"
        try {
            $groups = Invoke-WithThrottleProtection -OperationName "Get-PnPGroup" -ScriptBlock {
                Get-PnPGroup -ErrorAction Stop
            }

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
            }

            Write-ConsoleOutput "Retrieved $($importantGroups.Count) groups"
            Write-ConsoleOutput ""
        }
        catch {
            Write-ConsoleOutput "Failed to retrieve groups: $($_.Exception.Message)"
            Write-ConsoleOutput ""
        }

        # ===== ROLE ASSIGNMENT MAPPING =====
        Write-ConsoleOutput "Analyzing role assignments..."
        Update-Checkpoint -Phase "RoleAssignments"
        $raCounter = 0
        try {
            $web = Get-PnPWeb -ErrorAction Stop
            $roleAssignments = Invoke-WithThrottleProtection -OperationName "Get RoleAssignments" -ScriptBlock {
                Get-PnPProperty -ClientObject $web -Property RoleAssignments -ErrorAction Stop
            }

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
                                ScopeUrl = $SiteUrl
                                SiteTitle = $web.Title
                            }
                            Add-SharePointRoleAssignment -RoleData $roleData
                            $raCounter++
                        }
                    }
                }
                catch { }
            }

            Write-ConsoleOutput "Found $raCounter site-level role assignments"
            Write-ConsoleOutput ""
        }
        catch {
            Write-ConsoleOutput "Limited access to role assignment data: $($_.Exception.Message)"
            Write-ConsoleOutput ""
        }

        # ===== PERMISSION INHERITANCE TREE =====
        Write-ConsoleOutput "Checking permission inheritance..."
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

            Write-ConsoleOutput "Scanned $($visibleLists.Count) lists/libraries - $brokenCount with broken inheritance"
            Write-ConsoleOutput ""
        }
        catch {
            Write-ConsoleOutput "Limited access to inheritance data: $($_.Exception.Message)"
            Write-ConsoleOutput ""
        }

        # ===== SHARING LINKS AUDIT =====
        Write-ConsoleOutput "Auditing sharing links..."
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

            Write-ConsoleOutput "Found $linkCounter sharing links"
            Write-ConsoleOutput ""
        }
        catch {
            Write-ConsoleOutput "Limited access to sharing links data: $($_.Exception.Message)"
            Write-ConsoleOutput ""
        }

        # Show storage notice if we couldn't get it
        if ($storageValue -eq 0) {
            Write-ConsoleOutput "Note: Storage data unavailable. This requires SharePoint Administrator permissions or Sites.FullControl.All API permission."
            Write-ConsoleOutput ""
        }

        # Show analysis summary
        $p3Metrics = Get-SharePointData -DataType "Metrics"
        Write-ConsoleOutput "CORE ANALYSIS SUMMARY:"
        Write-ConsoleOutput "   Role Assignments: $($p3Metrics.TotalRoleAssignments)"
        Write-ConsoleOutput "   Inheritance Breaks: $($p3Metrics.InheritanceBreaks)"
        Write-ConsoleOutput "   Sharing Links: $($p3Metrics.TotalSharingLinks)"
        Write-ConsoleOutput ""
        Write-ConsoleOutput "PERMISSIONS ANALYSIS COMPLETED SUCCESSFULLY"

        Write-ActivityLog "Permissions analysis completed with storage: $storageValue MB" -Level "Information"
        Complete-Checkpoint -Status "Completed"
        Write-AuditEvent -EventType "DataCollection" -Detail "Permissions analysis complete" -AffectedObject $SiteUrl
        Complete-AuditSession -Status "Completed"

        # Log throttle stats if any retries occurred
        $throttleStats = Get-ThrottleStats
        if ($throttleStats.TotalRetries -gt 0) {
            Write-ConsoleOutput "Throttle protection: $($throttleStats.TotalRetries) retries, $($throttleStats.TotalThrottled) throttle events"
        }
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-RealPermissions-DataDriven"
        Complete-Checkpoint -Status "Failed"
        Write-AuditEvent -EventType "Error" -Detail $_.Exception.Message
        Complete-AuditSession -Status "Failed"
        Write-ConsoleOutput "ERROR: Permissions analysis failed"
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)"
    }
}
