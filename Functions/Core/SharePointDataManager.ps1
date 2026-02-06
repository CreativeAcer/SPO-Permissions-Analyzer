# ============================================
# SharePoint Data Manager
# ============================================
# This module manages SharePoint data in structured variables
# instead of parsing console text

# Global data storage
$script:SharePointData = @{
    Sites = @()
    Users = @()
    Groups = @()
    Permissions = @()
    RoleAssignments = @()
    InheritanceItems = @()
    SharingLinks = @()
    LastOperation = ""
    LastUpdateTime = $null
    OperationMetrics = @{
        TotalSites = 0
        TotalUsers = 0
        TotalGroups = 0
        ExternalUsers = 0
        SecurityFindings = 0
        RecordsProcessed = 0
        TotalRoleAssignments = 0
        InheritanceBreaks = 0
        TotalSharingLinks = 0
    }
}

function Initialize-SharePointDataManager {
    <#
    .SYNOPSIS
    Initializes the SharePoint data manager
    #>
    $script:SharePointData = @{
        Sites = @()
        Users = @()
        Groups = @()
        Permissions = @()
        RoleAssignments = @()
        InheritanceItems = @()
        SharingLinks = @()
        LastOperation = ""
        LastUpdateTime = $null
        OperationMetrics = @{
            TotalSites = 0
            TotalUsers = 0
            TotalGroups = 0
            ExternalUsers = 0
            SecurityFindings = 0
            RecordsProcessed = 0
            TotalRoleAssignments = 0
            InheritanceBreaks = 0
            TotalSharingLinks = 0
        }
    }

    Write-ActivityLog "SharePoint Data Manager initialized" -Level "Information"
}

function Add-SharePointSite {
    <#
    .SYNOPSIS
    Adds a site to the data store
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$SiteData
    )
    
    # Ensure required fields exist
    if (-not $SiteData.ContainsKey("Title")) { $SiteData["Title"] = "Unknown Site" }
    if (-not $SiteData.ContainsKey("Url")) { $SiteData["Url"] = "N/A" }
    if (-not $SiteData.ContainsKey("Owner")) { $SiteData["Owner"] = "N/A" }
    if (-not $SiteData.ContainsKey("Storage")) { $SiteData["Storage"] = "0" }
    if (-not $SiteData.ContainsKey("UsageLevel")) { 
        # Calculate usage level based on storage
        $storage = [int]$SiteData["Storage"]
        if ($storage -lt 500) {
            $SiteData["UsageLevel"] = "Low"
            $SiteData["UsageColor"] = "#28A745"
        } elseif ($storage -lt 1000) {
            $SiteData["UsageLevel"] = "Medium"
            $SiteData["UsageColor"] = "#FFC107"
        } elseif ($storage -lt 1500) {
            $SiteData["UsageLevel"] = "High"
            $SiteData["UsageColor"] = "#DC3545"
        } else {
            $SiteData["UsageLevel"] = "Critical"
            $SiteData["UsageColor"] = "#6F42C1"
        }
    }
    
    $script:SharePointData.Sites += $SiteData
    $script:SharePointData.OperationMetrics.TotalSites = $script:SharePointData.Sites.Count
    
    Write-ActivityLog "Added site: $($SiteData['Title'])" -Level "Information"
}

function Add-SharePointUser {
    <#
    .SYNOPSIS
    Adds a user to the data store
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$UserData
    )
    
    # Ensure required fields
    if (-not $UserData.ContainsKey("Name")) { $UserData["Name"] = "Unknown User" }
    if (-not $UserData.ContainsKey("Email")) { $UserData["Email"] = "N/A" }
    if (-not $UserData.ContainsKey("Type")) { $UserData["Type"] = "Internal" }
    if (-not $UserData.ContainsKey("Permission")) { $UserData["Permission"] = "Read" }
    
    $script:SharePointData.Users += $UserData
    $script:SharePointData.OperationMetrics.TotalUsers = $script:SharePointData.Users.Count
    
    # Count external users
    if ($UserData["Type"] -eq "External" -or $UserData["IsExternal"] -eq $true) {
        $script:SharePointData.OperationMetrics.ExternalUsers++
    }
    
    Write-ActivityLog "Added user: $($UserData['Name'])" -Level "Information"
}

function Add-SharePointGroup {
    <#
    .SYNOPSIS
    Adds a group to the data store
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$GroupData
    )
    
    # Ensure required fields
    if (-not $GroupData.ContainsKey("Name")) { $GroupData["Name"] = "Unknown Group" }
    if (-not $GroupData.ContainsKey("MemberCount")) { $GroupData["MemberCount"] = 0 }
    if (-not $GroupData.ContainsKey("Permission")) { $GroupData["Permission"] = "Read" }
    
    $script:SharePointData.Groups += $GroupData
    $script:SharePointData.OperationMetrics.TotalGroups = $script:SharePointData.Groups.Count
    
    Write-ActivityLog "Added group: $($GroupData['Name'])" -Level "Information"
}

function Add-SharePointRoleAssignment {
    <#
    .SYNOPSIS
    Adds a role assignment to the data store
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$RoleData
    )

    if (-not $RoleData.ContainsKey("Principal")) { $RoleData["Principal"] = "Unknown" }
    if (-not $RoleData.ContainsKey("PrincipalType")) { $RoleData["PrincipalType"] = "Unknown" }
    if (-not $RoleData.ContainsKey("Role")) { $RoleData["Role"] = "Unknown" }
    if (-not $RoleData.ContainsKey("Scope")) { $RoleData["Scope"] = "Site" }
    if (-not $RoleData.ContainsKey("ScopeUrl")) { $RoleData["ScopeUrl"] = "N/A" }

    $script:SharePointData.RoleAssignments += $RoleData
    $script:SharePointData.OperationMetrics.TotalRoleAssignments = $script:SharePointData.RoleAssignments.Count
}

function Add-SharePointInheritanceItem {
    <#
    .SYNOPSIS
    Adds an inheritance item to the data store
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$InheritanceData
    )

    if (-not $InheritanceData.ContainsKey("Title")) { $InheritanceData["Title"] = "Unknown" }
    if (-not $InheritanceData.ContainsKey("Url")) { $InheritanceData["Url"] = "N/A" }
    if (-not $InheritanceData.ContainsKey("Type")) { $InheritanceData["Type"] = "Unknown" }
    if (-not $InheritanceData.ContainsKey("HasUniquePermissions")) { $InheritanceData["HasUniquePermissions"] = $false }

    $script:SharePointData.InheritanceItems += $InheritanceData
    if ($InheritanceData["HasUniquePermissions"] -eq $true) {
        $script:SharePointData.OperationMetrics.InheritanceBreaks++
    }
}

function Add-SharePointSharingLink {
    <#
    .SYNOPSIS
    Adds a sharing link to the data store
    #>
    param(
        [Parameter(Mandatory=$true)]
        [hashtable]$LinkData
    )

    if (-not $LinkData.ContainsKey("GroupName")) { $LinkData["GroupName"] = "Unknown" }
    if (-not $LinkData.ContainsKey("LinkType")) { $LinkData["LinkType"] = "Unknown" }
    if (-not $LinkData.ContainsKey("AccessLevel")) { $LinkData["AccessLevel"] = "Unknown" }
    if (-not $LinkData.ContainsKey("MemberCount")) { $LinkData["MemberCount"] = 0 }

    $script:SharePointData.SharingLinks += $LinkData
    $script:SharePointData.OperationMetrics.TotalSharingLinks = $script:SharePointData.SharingLinks.Count
}

function Clear-SharePointData {
    <#
    .SYNOPSIS
    Clears all SharePoint data
    #>
    param(
        [string]$DataType = "All"
    )

    switch ($DataType) {
        "Sites" {
            $script:SharePointData.Sites = @()
            $script:SharePointData.OperationMetrics.TotalSites = 0
        }
        "Users" {
            $script:SharePointData.Users = @()
            $script:SharePointData.OperationMetrics.TotalUsers = 0
            $script:SharePointData.OperationMetrics.ExternalUsers = 0
        }
        "Groups" {
            $script:SharePointData.Groups = @()
            $script:SharePointData.OperationMetrics.TotalGroups = 0
        }
        "RoleAssignments" {
            $script:SharePointData.RoleAssignments = @()
            $script:SharePointData.OperationMetrics.TotalRoleAssignments = 0
        }
        "InheritanceItems" {
            $script:SharePointData.InheritanceItems = @()
            $script:SharePointData.OperationMetrics.InheritanceBreaks = 0
        }
        "SharingLinks" {
            $script:SharePointData.SharingLinks = @()
            $script:SharePointData.OperationMetrics.TotalSharingLinks = 0
        }
        "All" {
            Initialize-SharePointDataManager
        }
    }

    Write-ActivityLog "Cleared SharePoint data: $DataType" -Level "Information"
}

function Get-SharePointData {
    <#
    .SYNOPSIS
    Gets SharePoint data from the store
    #>
    param(
        [string]$DataType = "All"
    )
    
    switch ($DataType) {
        "Sites" { return $script:SharePointData.Sites }
        "Users" { return $script:SharePointData.Users }
        "Groups" { return $script:SharePointData.Groups }
        "Permissions" { return $script:SharePointData.Permissions }
        "RoleAssignments" { return $script:SharePointData.RoleAssignments }
        "InheritanceItems" { return $script:SharePointData.InheritanceItems }
        "SharingLinks" { return $script:SharePointData.SharingLinks }
        "Metrics" { return $script:SharePointData.OperationMetrics }
        "All" { return $script:SharePointData }
    }
}

function Set-SharePointOperationContext {
    <#
    .SYNOPSIS
    Sets the current operation context
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$OperationType
    )
    
    $script:SharePointData.LastOperation = $OperationType
    $script:SharePointData.LastUpdateTime = Get-Date
    
    Write-ActivityLog "Set operation context: $OperationType" -Level "Information"
}

function Update-VisualAnalyticsFromData {
    <#
    .SYNOPSIS
    Updates Visual Analytics directly from stored data
    #>
    try {
        Write-ActivityLog "=== UPDATING VISUAL ANALYTICS FROM DATA ===" -Level "Information"
        
        # Get current data
        $metrics = Get-SharePointData -DataType "Metrics"
        $sites = Get-SharePointData -DataType "Sites"
        
        Write-ActivityLog "Data summary: Sites=$($sites.Count), Users=$($metrics.TotalUsers), Groups=$($metrics.TotalGroups), External=$($metrics.ExternalUsers)" -Level "Information"
        
        # Update metrics cards
        if ($script:txtTotalSites) { $script:txtTotalSites.Text = $metrics.TotalSites.ToString() }
        if ($script:txtTotalUsers) { $script:txtTotalUsers.Text = $metrics.TotalUsers.ToString() }
        if ($script:txtTotalGroups) { $script:txtTotalGroups.Text = $metrics.TotalGroups.ToString() }
        if ($script:txtExternalUsers) { $script:txtExternalUsers.Text = $metrics.ExternalUsers.ToString() }

        # Update security analysis cards (P3)
        if ($script:txtRoleAssignments) { $script:txtRoleAssignments.Text = $metrics.TotalRoleAssignments.ToString() }
        if ($script:txtInheritanceBreaks) { $script:txtInheritanceBreaks.Text = $metrics.InheritanceBreaks.ToString() }
        if ($script:txtSharingLinks) { $script:txtSharingLinks.Text = $metrics.TotalSharingLinks.ToString() }
        
        # Update sites data grid
        if ($sites.Count -gt 0 -and $script:dgSites) {
            $siteObjects = @()
            foreach ($site in $sites) {
                $siteObjects += [PSCustomObject]@{
                    Title = $site["Title"]
                    Url = $site["Url"]
                    Owner = $site["Owner"]
                    Storage = "$($site['Storage']) MB"
                    UsageLevel = $site["UsageLevel"]
                    UsageColor = $site["UsageColor"]
                    UserCount = if ($site["UserCount"]) { $site["UserCount"] } else { 0 }
                    GroupCount = if ($site["GroupCount"]) { $site["GroupCount"] } else { 0 }
                }
            }
            $script:dgSites.ItemsSource = $siteObjects
            Write-ActivityLog "Updated data grid with $($siteObjects.Count) sites" -Level "Information"
        }
        
        # Update charts
        if ($script:canvasStorageChart -and $script:canvasPermissionChart) {
            Update-Charts -SitesData $sites
            Write-ActivityLog "Updated charts" -Level "Information"
        }
        
        # Generate alerts
        if ($script:lstPermissionAlerts) {
            Generate-AnalyticsAlerts `
                -SitesCount $metrics.TotalSites `
                -UsersCount $metrics.TotalUsers `
                -ExternalCount $metrics.ExternalUsers `
                -GroupsCount $metrics.TotalGroups `
                -RecordsProcessed $metrics.RecordsProcessed `
                -SecurityFindings $metrics.SecurityFindings
            Write-ActivityLog "Generated alerts" -Level "Information"
        }
        
        # Update subtitle
        if ($script:txtAnalyticsSubtitle) {
            $operation = $script:SharePointData.LastOperation
            if (-not $operation) { $operation = "Analysis" }
            $script:txtAnalyticsSubtitle.Text = "Last updated: $(Get-Date -Format 'MMM dd, yyyy HH:mm:ss') - Data from $operation"
        }
        
        Write-ActivityLog "=== VISUAL ANALYTICS UPDATE COMPLETED ===" -Level "Information"
        
    }
    catch {
        Write-ActivityLog "ERROR updating Visual Analytics from data: $($_.Exception.Message)" -Level "Error"
        Write-ActivityLog "Stack trace: $($_.Exception.StackTrace)" -Level "Error"
    }
}