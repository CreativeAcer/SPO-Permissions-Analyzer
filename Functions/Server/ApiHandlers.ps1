# ============================================
# ApiHandlers.ps1 - REST API Endpoint Handlers
# ============================================
# Bridges the web frontend to existing PowerShell functions.
# All data collection reuses SharePointDataManager & SPOConnection.

function Invoke-ApiHandler {
    <#
    .SYNOPSIS
    Routes API requests to the correct handler
    #>
    param(
        [System.Net.HttpListenerRequest]$Request,
        [System.Net.HttpListenerResponse]$Response,
        [string]$Path
    )

    $method = $Request.HttpMethod

    switch -Wildcard ($Path) {
        "/api/status"       { Handle-GetStatus -Response $Response }
        "/api/connect"      { Handle-PostConnect -Request $Request -Response $Response }
        "/api/demo"         { Handle-PostDemo -Response $Response }
        "/api/sites"        { Handle-PostSites -Request $Request -Response $Response }
        "/api/permissions"  { Handle-PostPermissions -Request $Request -Response $Response }
        "/api/progress"     { Handle-GetProgress -Response $Response }
        "/api/data/*"       {
            $dataType = $Path.Replace("/api/data/", "")
            Handle-GetData -Response $Response -DataType $dataType
        }
        "/api/metrics"      { Handle-GetMetrics -Response $Response }
        "/api/export/*"     {
            $exportType = $Path.Replace("/api/export/", "")
            Handle-PostExport -Request $Request -Response $Response -ExportType $exportType
        }
        "/api/export-json"  { Handle-PostExportJson -Response $Response }
        "/api/export-json/*" {
            $jsonType = $Path.Replace("/api/export-json/", "")
            Handle-PostExportJsonType -Response $Response -DataType $jsonType
        }
        "/api/shutdown"     {
            Send-JsonResponse -Response $Response -Data @{ success = $true; message = "Shutting down" }
            Stop-WebServer
        }
        default {
            Send-JsonResponse -Response $Response -Data @{ error = "Unknown endpoint: $Path" } -StatusCode 404
        }
    }
}

# ---- Status ----

function Handle-GetStatus {
    param($Response)

    $metrics = Get-SharePointData -DataType "Metrics"

    Send-JsonResponse -Response $Response -Data @{
        connected = [bool]$script:SPOConnected
        demoMode  = [bool]$script:DemoMode
        headless  = [bool]$env:SPO_HEADLESS
        lastOperation = $script:SharePointData.LastOperation
        metrics = @{
            totalSites = $metrics.TotalSites
            totalUsers = $metrics.TotalUsers
            totalGroups = $metrics.TotalGroups
            externalUsers = $metrics.ExternalUsers
            totalRoleAssignments = $metrics.TotalRoleAssignments
            inheritanceBreaks = $metrics.InheritanceBreaks
            totalSharingLinks = $metrics.TotalSharingLinks
        }
    }
}

# ---- Connect ----

function Handle-PostConnect {
    param($Request, $Response)

    $body = Read-RequestBody -Request $Request
    if (-not $body -or -not $body.tenantUrl -or -not $body.clientId) {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = "tenantUrl and clientId are required" } -StatusCode 400
        return
    }

    try {
        # Save settings
        Set-AppSetting -SettingName "SharePoint.TenantUrl" -Value $body.tenantUrl
        Set-AppSetting -SettingName "SharePoint.ClientId" -Value $body.clientId

        # Check PnP module
        if (-not (Test-PnPModuleAvailable)) {
            Send-JsonResponse -Response $Response -Data @{
                success = $false
                message = "PnP.PowerShell module not found. Please install it first."
            }
            return
        }

        # Attempt connection
        Write-ActivityLog "Web UI connecting to: $($body.tenantUrl)" -Level "Information"

        if ($env:SPO_HEADLESS) {
            # Container/headless mode: use device code flow
            # The device code appears in the container terminal (podman logs / docker logs)
            Write-Host ""
            Write-Host "  Device code authentication requested from Web UI" -ForegroundColor Yellow
            Write-Host "  Tenant: $($body.tenantUrl)" -ForegroundColor White
            Write-Host ""
            Connect-PnPOnline -Url $body.tenantUrl -ClientId $body.clientId -DeviceLogin
        }
        else {
            # Host mode: use interactive browser popup
            Connect-PnPOnline -Url $body.tenantUrl -ClientId $body.clientId -Interactive
        }

        $web = Get-PnPWeb -ErrorAction SilentlyContinue
        $script:SPOConnected = $true

        # Get current user
        $currentUser = "Connected"
        try {
            $userInfo = Get-PnPCurrentUser -ErrorAction SilentlyContinue
            if ($userInfo) {
                $currentUser = @($userInfo.UserPrincipalName, $userInfo.Email, $userInfo.LoginName) |
                    Where-Object { $_ } | Select-Object -First 1
                if (-not $currentUser) { $currentUser = "Connected" }
            }
        }
        catch { }

        Send-JsonResponse -Response $Response -Data @{
            success = $true
            message = "Connected to SharePoint Online"
            siteTitle = if ($web) { $web.Title } else { "SharePoint Site" }
            siteUrl = if ($web) { $web.Url } else { $body.tenantUrl }
            user = $currentUser
        }
    }
    catch {
        $script:SPOConnected = $false
        Send-JsonResponse -Response $Response -Data @{
            success = $false
            message = "Connection failed: $($_.Exception.Message)"
        }
    }
}

# ---- Demo Mode ----

function Handle-PostDemo {
    param($Response)

    try {
        $script:DemoMode = $true
        $script:SPOConnected = $true

        Initialize-SharePointDataManager

        # Re-use existing demo data generators
        $script:ServerState.OperationLog.Clear()
        $script:ServerState.OperationRunning = $true
        $script:ServerState.OperationComplete = $false

        Add-OperationLog "Starting Demo Mode..."
        Add-OperationLog "Generating simulated SharePoint data..."

        # Generate demo sites
        Set-SharePointOperationContext -OperationType "Demo - Sites"
        $demoSites = @(
            @{Title="Team Collaboration Site"; Url="https://contoso.sharepoint.com/sites/teamsite"; Owner="admin@contoso.com"; Storage="750"; Template="STS#3"; LastModified="2025-01-15"; UserCount=25; GroupCount=8},
            @{Title="HR Portal"; Url="https://contoso.sharepoint.com/sites/hr"; Owner="hr.admin@contoso.com"; Storage="1200"; Template="SITEPAGEPUBLISHING#0"; LastModified="2025-02-01"; UserCount=45; GroupCount=12},
            @{Title="Marketing Hub"; Url="https://contoso.sharepoint.com/sites/marketing"; Owner="marketing@contoso.com"; Storage="2500"; Template="STS#3"; LastModified="2025-01-28"; UserCount=30; GroupCount=6},
            @{Title="Finance Department"; Url="https://contoso.sharepoint.com/sites/finance"; Owner="finance.lead@contoso.com"; Storage="450"; Template="STS#3"; LastModified="2025-01-10"; UserCount=15; GroupCount=4},
            @{Title="Executive Dashboard"; Url="https://contoso.sharepoint.com/sites/exec"; Owner="ceo@contoso.com"; Storage="200"; Template="SITEPAGEPUBLISHING#0"; LastModified="2024-12-15"; UserCount=8; GroupCount=3}
        )
        foreach ($site in $demoSites) { Add-SharePointSite -SiteData $site }
        Add-OperationLog "Added $($demoSites.Count) demo sites"

        # Generate demo users
        Set-SharePointOperationContext -OperationType "Demo - Permissions"
        $demoUsers = @(
            @{Name="John Doe"; Email="john.doe@contoso.com"; Type="Internal"; Permission="Full Control"; IsSiteAdmin=$true; LoginName="i:0#.f|membership|john.doe@contoso.com"},
            @{Name="Jane Smith"; Email="jane.smith@contoso.com"; Type="Internal"; Permission="Edit"; IsSiteAdmin=$false},
            @{Name="Mike Johnson"; Email="mike.j@contoso.com"; Type="Internal"; Permission="Read"; IsSiteAdmin=$false},
            @{Name="Sarah Wilson"; Email="sarah.w@contoso.com"; Type="Internal"; Permission="Contribute"; IsSiteAdmin=$false},
            @{Name="External Partner"; Email="partner@external.com"; Type="External"; Permission="Read"; IsExternal=$true},
            @{Name="Guest Reviewer"; Email="reviewer@partner.org"; Type="External"; Permission="Read"; IsExternal=$true},
            @{Name="Contractor A"; Email="contractor.a@vendor.com"; Type="External"; Permission="Edit"; IsExternal=$true},
            @{Name="David Brown"; Email="david.b@contoso.com"; Type="Internal"; Permission="Full Control"; IsSiteAdmin=$true},
            @{Name="Emily Chen"; Email="emily.c@contoso.com"; Type="Internal"; Permission="Edit"; IsSiteAdmin=$false},
            @{Name="Alex Kumar"; Email="alex.k@contoso.com"; Type="Internal"; Permission="Read"; IsSiteAdmin=$false}
        )
        foreach ($user in $demoUsers) { Add-SharePointUser -UserData $user }
        Add-OperationLog "Added $($demoUsers.Count) demo users"

        # Generate demo groups
        $demoGroups = @(
            @{Name="Site Owners"; MemberCount=3; Permission="Full Control"; Description="Owners of the site"},
            @{Name="Site Members"; MemberCount=12; Permission="Edit"; Description="Members with edit access"},
            @{Name="Site Visitors"; MemberCount=25; Permission="Read"; Description="Visitors with read access"},
            @{Name="HR Team"; MemberCount=8; Permission="Contribute"; Description="Human Resources team"},
            @{Name="IT Admins"; MemberCount=4; Permission="Full Control"; Description="IT administrators"},
            @{Name="Marketing Team"; MemberCount=15; Permission="Edit"; Description="Marketing department"}
        )
        foreach ($group in $demoGroups) { Add-SharePointGroup -GroupData $group }
        Add-OperationLog "Added $($demoGroups.Count) demo groups"

        # Generate demo role assignments
        $demoRoles = @(
            @{Principal="John Doe"; PrincipalType="User"; Role="Full Control"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
            @{Principal="Jane Smith"; PrincipalType="User"; Role="Edit"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
            @{Principal="Mike Johnson"; PrincipalType="User"; Role="Read"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
            @{Principal="External Partner"; PrincipalType="User"; Role="Read"; Scope="Library"; ScopeUrl="/sites/teamsite/Shared Documents"; SiteTitle="Team Collaboration Site"},
            @{Principal="Site Owners"; PrincipalType="SharePoint Group"; Role="Full Control"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
            @{Principal="Site Members"; PrincipalType="SharePoint Group"; Role="Edit"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
            @{Principal="Site Visitors"; PrincipalType="SharePoint Group"; Role="Read"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
            @{Principal="HR Team"; PrincipalType="Security Group"; Role="Contribute"; Scope="Library"; ScopeUrl="/sites/hr/Shared Documents"; SiteTitle="HR Portal"},
            @{Principal="Finance Group"; PrincipalType="Security Group"; Role="Read"; Scope="List"; ScopeUrl="/sites/finance/Lists/Budget"; SiteTitle="Finance Department"},
            @{Principal="IT Admins"; PrincipalType="Security Group"; Role="Full Control"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
            @{Principal="Marketing Team"; PrincipalType="Security Group"; Role="Edit"; Scope="Library"; ScopeUrl="/sites/marketing/Assets"; SiteTitle="Marketing Hub"},
            @{Principal="Contractor A"; PrincipalType="User"; Role="Edit"; Scope="Library"; ScopeUrl="/sites/teamsite/Project Files"; SiteTitle="Team Collaboration Site"}
        )
        foreach ($ra in $demoRoles) { Add-SharePointRoleAssignment -RoleData $ra }
        Add-OperationLog "Added $($demoRoles.Count) role assignments"

        # Generate demo inheritance items
        $demoInheritance = @(
            @{Title="Team Collaboration Site"; Type="Site"; Url="https://contoso.sharepoint.com/sites/teamsite"; HasUniquePermissions=$true; ParentUrl="N/A"; RoleAssignmentCount=7; SiteTitle="Team Collaboration Site"},
            @{Title="Shared Documents"; Type="Document Library"; Url="/sites/teamsite/Shared Documents"; HasUniquePermissions=$true; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=4; SiteTitle="Team Collaboration Site"},
            @{Title="Site Pages"; Type="Document Library"; Url="/sites/teamsite/SitePages"; HasUniquePermissions=$false; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"},
            @{Title="Project Files"; Type="Document Library"; Url="/sites/teamsite/Project Files"; HasUniquePermissions=$true; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=5; SiteTitle="Team Collaboration Site"},
            @{Title="Marketing Assets"; Type="Document Library"; Url="/sites/marketing/Assets"; HasUniquePermissions=$true; ParentUrl="https://contoso.sharepoint.com/sites/marketing"; RoleAssignmentCount=3; SiteTitle="Marketing Hub"},
            @{Title="Budget Tracker"; Type="List"; Url="/sites/finance/Lists/Budget"; HasUniquePermissions=$true; ParentUrl="https://contoso.sharepoint.com/sites/finance"; RoleAssignmentCount=2; SiteTitle="Finance Department"},
            @{Title="Team Calendar"; Type="List"; Url="/sites/teamsite/Lists/Calendar"; HasUniquePermissions=$false; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"},
            @{Title="Announcements"; Type="List"; Url="/sites/teamsite/Lists/Announcements"; HasUniquePermissions=$false; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"},
            @{Title="HR Policies"; Type="Document Library"; Url="/sites/hr/Policies"; HasUniquePermissions=$true; ParentUrl="https://contoso.sharepoint.com/sites/hr"; RoleAssignmentCount=3; SiteTitle="HR Portal"},
            @{Title="Style Library"; Type="Document Library"; Url="/sites/teamsite/Style Library"; HasUniquePermissions=$false; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"}
        )
        foreach ($item in $demoInheritance) { Add-SharePointInheritanceItem -InheritanceData $item }
        Add-OperationLog "Added $($demoInheritance.Count) inheritance items"

        # Generate demo sharing links
        $demoLinks = @(
            @{GroupName="SharingLinks.abc123.OrganizationView.def456"; LinkType="Company-wide"; AccessLevel="View"; MemberCount=0; SiteTitle="Team Collaboration Site"; CreatedDate="2025-01-15"},
            @{GroupName="SharingLinks.abc124.OrganizationEdit.ghi789"; LinkType="Company-wide"; AccessLevel="Edit"; MemberCount=0; SiteTitle="Marketing Hub"; CreatedDate="2025-02-01"},
            @{GroupName="SharingLinks.abc125.AnonymousView.jkl012"; LinkType="Anonymous"; AccessLevel="View"; MemberCount=3; SiteTitle="Marketing Hub"; CreatedDate="2025-01-20"},
            @{GroupName="SharingLinks.abc126.Flexible.mno345"; LinkType="Specific People"; AccessLevel="Edit"; MemberCount=5; SiteTitle="Team Collaboration Site"; CreatedDate="2025-02-10"},
            @{GroupName="SharingLinks.abc127.Flexible.pqr678"; LinkType="Specific People"; AccessLevel="View"; MemberCount=2; SiteTitle="HR Portal"; CreatedDate="2025-01-25"},
            @{GroupName="SharingLinks.abc128.AnonymousEdit.stu901"; LinkType="Anonymous"; AccessLevel="Edit"; MemberCount=1; SiteTitle="Team Collaboration Site"; CreatedDate="2025-03-01"},
            @{GroupName="SharingLinks.abc129.Flexible.vwx234"; LinkType="Specific People"; AccessLevel="Edit"; MemberCount=8; SiteTitle="Finance Department"; CreatedDate="2025-02-15"}
        )
        foreach ($link in $demoLinks) { Add-SharePointSharingLink -LinkData $link }
        Add-OperationLog "Added $($demoLinks.Count) sharing links"

        $metrics = Get-SharePointData -DataType "Metrics"
        Add-OperationLog ""
        Add-OperationLog "Demo mode ready!"
        Add-OperationLog "Sites: $($metrics.TotalSites) | Users: $($metrics.TotalUsers) | Groups: $($metrics.TotalGroups)"
        Add-OperationLog "External: $($metrics.ExternalUsers) | Roles: $($metrics.TotalRoleAssignments) | Inheritance Breaks: $($metrics.InheritanceBreaks) | Links: $($metrics.TotalSharingLinks)"

        $script:ServerState.OperationRunning = $false
        $script:ServerState.OperationComplete = $true

        Send-JsonResponse -Response $Response -Data @{
            success = $true
            message = "Demo mode activated with sample data"
        }
    }
    catch {
        $script:ServerState.OperationRunning = $false
        $script:ServerState.OperationComplete = $true
        Send-JsonResponse -Response $Response -Data @{
            success = $false
            message = "Demo mode failed: $($_.Exception.Message)"
        }
    }
}

# ---- Sites ----

function Handle-PostSites {
    param($Request, $Response)

    if (-not $script:SPOConnected -and -not $script:DemoMode) {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = "Not connected" } -StatusCode 400
        return
    }

    try {
        if ($script:DemoMode) {
            Send-JsonResponse -Response $Response -Data @{ success = $true; message = "Sites already loaded in demo mode" }
            return
        }

        # Use existing real sites function logic
        $script:ServerState.OperationLog.Clear()
        Add-OperationLog "Fetching sites..."
        Get-RealSites-DataDriven
        Add-OperationLog "Sites loaded successfully"

        Send-JsonResponse -Response $Response -Data @{ success = $true; message = "Sites retrieved" }
    }
    catch {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = $_.Exception.Message }
    }
}

# ---- Permissions Analysis ----

function Handle-PostPermissions {
    param($Request, $Response)

    if (-not $script:SPOConnected -and -not $script:DemoMode) {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = "Not connected" } -StatusCode 400
        return
    }

    $body = Read-RequestBody -Request $Request
    $siteUrl = if ($body -and $body.siteUrl) { $body.siteUrl } else { "" }

    try {
        $script:ServerState.OperationLog.Clear()
        $script:ServerState.OperationRunning = $true
        $script:ServerState.OperationComplete = $false

        if ($script:DemoMode) {
            Add-OperationLog "Permissions data already loaded in demo mode."
            $script:ServerState.OperationRunning = $false
            $script:ServerState.OperationComplete = $true
            Send-JsonResponse -Response $Response -Data @{ success = $true; message = "Permissions available from demo data" }
            return
        }

        # Run real analysis - this is synchronous and may take a while
        Add-OperationLog "Starting permissions analysis..."
        Get-RealPermissions-DataDriven -SiteUrl $siteUrl
        Add-OperationLog "Analysis complete."

        $script:ServerState.OperationRunning = $false
        $script:ServerState.OperationComplete = $true

        Send-JsonResponse -Response $Response -Data @{ success = $true; message = "Permissions analysis complete" }
    }
    catch {
        $script:ServerState.OperationRunning = $false
        $script:ServerState.OperationComplete = $true
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = $_.Exception.Message }
    }
}

# ---- Progress ----

function Handle-GetProgress {
    param($Response)

    Send-JsonResponse -Response $Response -Data @{
        messages = @($script:ServerState.OperationLog.ToArray())
        running = $script:ServerState.OperationRunning
        complete = $script:ServerState.OperationComplete
    }
}

# ---- Data ----

function Handle-GetData {
    param($Response, [string]$DataType)

    $typeMap = @{
        "sites"           = "Sites"
        "users"           = "Users"
        "groups"          = "Groups"
        "permissions"     = "Permissions"
        "roleassignments" = "RoleAssignments"
        "inheritance"     = "InheritanceItems"
        "sharinglinks"    = "SharingLinks"
    }

    $mappedType = if ($typeMap.ContainsKey($DataType.ToLower())) { $typeMap[$DataType.ToLower()] } else { $DataType }

    $data = Get-SharePointData -DataType $mappedType
    if ($null -eq $data) { $data = @() }

    Send-JsonResponse -Response $Response -Data @{ data = @($data) }
}

# ---- Metrics ----

function Handle-GetMetrics {
    param($Response)

    $metrics = Get-SharePointData -DataType "Metrics"

    Send-JsonResponse -Response $Response -Data @{
        totalSites = $metrics.TotalSites
        totalUsers = $metrics.TotalUsers
        totalGroups = $metrics.TotalGroups
        externalUsers = $metrics.ExternalUsers
        securityFindings = $metrics.SecurityFindings
        totalRoleAssignments = $metrics.TotalRoleAssignments
        inheritanceBreaks = $metrics.InheritanceBreaks
        totalSharingLinks = $metrics.TotalSharingLinks
    }
}

# ---- JSON Export ----

function Handle-PostExportJson {
    param($Response)

    try {
        $report = Build-GovernanceReport -IncludeMetadata
        Send-JsonResponse -Response $Response -Data $report
    }
    catch {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = $_.Exception.Message }
    }
}

function Handle-PostExportJsonType {
    param($Response, [string]$DataType)

    try {
        $typeMap = @{
            "sites"           = "Sites"
            "users"           = "Users"
            "groups"          = "Groups"
            "roleassignments" = "RoleAssignments"
            "inheritance"     = "InheritanceItems"
            "sharinglinks"    = "SharingLinks"
        }

        $mappedType = if ($typeMap.ContainsKey($DataType.ToLower())) { $typeMap[$DataType.ToLower()] } else { $DataType }
        $data = Get-SharePointData -DataType $mappedType

        $output = [ordered]@{
            schemaVersion = "1.0.0"
            exportedAt    = (Get-Date).ToString("o")
            dataType      = $DataType
            count         = @($data).Count
            data          = @($data)
        }

        Send-JsonResponse -Response $Response -Data $output
    }
    catch {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = $_.Exception.Message }
    }
}

# ---- Export (CSV) ----

function Handle-PostExport {
    param($Request, $Response, [string]$ExportType)

    try {
        $typeMap = @{
            "sites"           = "Sites"
            "users"           = "Users"
            "groups"          = "Groups"
            "roleassignments" = "RoleAssignments"
            "inheritance"     = "InheritanceItems"
            "sharinglinks"    = "SharingLinks"
        }

        $mappedType = if ($typeMap.ContainsKey($ExportType.ToLower())) { $typeMap[$ExportType.ToLower()] } else { $ExportType }
        $data = Get-SharePointData -DataType $mappedType

        if (-not $data -or $data.Count -eq 0) {
            Send-JsonResponse -Response $Response -Data @{ success = $false; message = "No data to export" }
            return
        }

        # Convert hashtable array to CSV-friendly objects
        $csvRows = @()
        foreach ($item in $data) {
            if ($item -is [hashtable]) {
                $csvRows += [PSCustomObject]$item
            } else {
                $csvRows += $item
            }
        }

        $csvContent = ($csvRows | ConvertTo-Csv -NoTypeInformation) -join "`n"

        # Send as downloadable CSV
        $Response.StatusCode = 200
        $Response.ContentType = "text/csv; charset=utf-8"
        $Response.Headers.Add("Content-Disposition", "attachment; filename=SharePoint_${ExportType}_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv")

        $bytes = [System.Text.Encoding]::UTF8.GetBytes($csvContent)
        $Response.ContentLength64 = $bytes.Length
        $Response.OutputStream.Write($bytes, 0, $bytes.Length)
        $Response.OutputStream.Close()
    }
    catch {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = $_.Exception.Message }
    }
}
