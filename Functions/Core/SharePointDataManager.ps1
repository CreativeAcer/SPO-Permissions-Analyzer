# ============================================
# SharePoint Data Manager
# ============================================
# This module manages SharePoint data in structured variables
# instead of parsing console text

# Global data storage
$script:SharePointData = @{
    Sites = [System.Collections.ArrayList]::new()
    Users = [System.Collections.ArrayList]::new()
    Groups = [System.Collections.ArrayList]::new()
    Permissions = [System.Collections.ArrayList]::new()
    RoleAssignments = [System.Collections.ArrayList]::new()
    InheritanceItems = [System.Collections.ArrayList]::new()
    SharingLinks = [System.Collections.ArrayList]::new()
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
        Sites = [System.Collections.ArrayList]::new()
        Users = [System.Collections.ArrayList]::new()
        Groups = [System.Collections.ArrayList]::new()
        Permissions = [System.Collections.ArrayList]::new()
        RoleAssignments = [System.Collections.ArrayList]::new()
        InheritanceItems = [System.Collections.ArrayList]::new()
        SharingLinks = [System.Collections.ArrayList]::new()
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
    
    [void]$script:SharePointData.Sites.Add($SiteData)
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
    
    [void]$script:SharePointData.Users.Add($UserData)
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
    
    [void]$script:SharePointData.Groups.Add($GroupData)
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

    [void]$script:SharePointData.RoleAssignments.Add($RoleData)
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

    [void]$script:SharePointData.InheritanceItems.Add($InheritanceData)
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

    [void]$script:SharePointData.SharingLinks.Add($LinkData)
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
            $script:SharePointData.Sites.Clear()
            $script:SharePointData.OperationMetrics.TotalSites = 0
        }
        "Users" {
            $script:SharePointData.Users.Clear()
            $script:SharePointData.OperationMetrics.TotalUsers = 0
            $script:SharePointData.OperationMetrics.ExternalUsers = 0
        }
        "Groups" {
            $script:SharePointData.Groups.Clear()
            $script:SharePointData.OperationMetrics.TotalGroups = 0
        }
        "RoleAssignments" {
            $script:SharePointData.RoleAssignments.Clear()
            $script:SharePointData.OperationMetrics.TotalRoleAssignments = 0
        }
        "InheritanceItems" {
            $script:SharePointData.InheritanceItems.Clear()
            $script:SharePointData.OperationMetrics.InheritanceBreaks = 0
        }
        "SharingLinks" {
            $script:SharePointData.SharingLinks.Clear()
            $script:SharePointData.OperationMetrics.TotalSharingLinks = 0
        }
        "All" {
            # Clear each collection in-place to preserve the shared hashtable reference
            # (Initialize-SharePointDataManager creates a new hashtable which breaks
            # the reference held by ServerState.SharePointData in background runspaces)
            $script:SharePointData.Sites.Clear()
            $script:SharePointData.Users.Clear()
            $script:SharePointData.Groups.Clear()
            $script:SharePointData.Permissions.Clear()
            $script:SharePointData.RoleAssignments.Clear()
            $script:SharePointData.InheritanceItems.Clear()
            $script:SharePointData.SharingLinks.Clear()
            $script:SharePointData.LastOperation = ""
            $script:SharePointData.LastUpdateTime = $null
            $script:SharePointData.OperationMetrics.TotalSites = 0
            $script:SharePointData.OperationMetrics.TotalUsers = 0
            $script:SharePointData.OperationMetrics.TotalGroups = 0
            $script:SharePointData.OperationMetrics.ExternalUsers = 0
            $script:SharePointData.OperationMetrics.SecurityFindings = 0
            $script:SharePointData.OperationMetrics.RecordsProcessed = 0
            $script:SharePointData.OperationMetrics.TotalRoleAssignments = 0
            $script:SharePointData.OperationMetrics.InheritanceBreaks = 0
            $script:SharePointData.OperationMetrics.TotalSharingLinks = 0
            Write-ActivityLog "SharePoint Data Manager cleared" -Level "Information"
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

