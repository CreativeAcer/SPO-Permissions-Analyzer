# ============================================
# PermissionsCollector.ps1 - SharePoint Permissions Analysis
# ============================================
# Analyzes real SharePoint permissions and stores in data manager.
# Runs inside background runspaces (Start-BackgroundOperation)
# where Write-ConsoleOutput writes to the shared operation log.

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
