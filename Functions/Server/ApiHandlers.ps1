# ============================================
# ApiHandlers.ps1 - REST API Endpoint Handlers
# ============================================
# Bridges the web frontend to existing PowerShell functions.
# All data collection reuses SharePointDataManager & SPOConnection.
# Demo data generation is delegated to Demo/DemoDataGenerator.ps1.

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
        "/api/permissions"        { Handle-PostPermissions -Request $Request -Response $Response }
        "/api/prepare-analysis"  { Handle-PostPrepareAnalysis -Request $Request -Response $Response }
        "/api/progress"     { Handle-GetProgress -Response $Response }
        "/api/data/*"       {
            $dataType = $Path.Replace("/api/data/", "")
            Handle-GetData -Response $Response -DataType $dataType
        }
        "/api/metrics"      { Handle-GetMetrics -Response $Response }
        "/api/enrich"       { Handle-PostEnrich -Response $Response }
        "/api/enrichment"   { Handle-GetEnrichment -Response $Response }
        "/api/risk"         { Handle-GetRisk -Response $Response }
        "/api/export/*"     {
            $exportType = $Path.Replace("/api/export/", "")
            Handle-PostExport -Request $Request -Response $Response -ExportType $exportType
        }
        "/api/export-json"  { Handle-PostExportJson -Response $Response }
        "/api/export-json/*" {
            $jsonType = $Path.Replace("/api/export-json/", "")
            Handle-PostExportJsonType -Response $Response -DataType $jsonType
        }
        "/api/audit"        { Handle-GetAudit -Response $Response }
        "/api/build-permissions-matrix" { Handle-PostBuildPermissionsMatrix -Request $Request -Response $Response }
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

        # Extract tenant name from URL for DeviceLogin authentication
        # e.g., https://contoso.sharepoint.com or https://contoso-admin.sharepoint.com -> contoso.onmicrosoft.com
        $tenantName = $null
        if ($body.tenantUrl -match '//([^-\.]+)') {
            # Extract base tenant name (before any hyphens like -admin or -my)
            $tenantName = "$($matches[1]).onmicrosoft.com"
        }

        # For DeviceLogin (container), use admin URL to enable tenant-wide operations
        # Interactive mode can auto-switch, but DeviceLogin cannot
        $connectionUrl = $body.tenantUrl
        if ($env:SPO_HEADLESS) {
            # Convert tenant root URL to admin URL for DeviceLogin (for admin capabilities)
            # But keep specific site URLs as-is
            if ($connectionUrl -notmatch '-admin\.sharepoint\.com') {
                # Check if this is a tenant root URL (no /sites/ or /teams/ path)
                if ($connectionUrl -match '^https://[^/]+\.sharepoint\.com/?$') {
                    # Tenant root - convert to admin for capability testing
                    $connectionUrl = $connectionUrl -replace '(https://[^\.]+)\.sharepoint\.com', '$1-admin.sharepoint.com'
                    Write-ActivityLog "Using admin URL for capability testing: $connectionUrl" -Level "Information"
                }
                # else: specific site URL - use as-is
            }
        }

        if ($env:SPO_HEADLESS) {
            # Container/headless mode: use device code flow
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "  DEVICE CODE AUTHENTICATION" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "  Requested URL: $($body.tenantUrl)" -ForegroundColor White
            Write-Host "  Connection URL: $connectionUrl" -ForegroundColor Yellow
            if ($tenantName) {
                Write-Host "  Tenant ID: $tenantName" -ForegroundColor White
            }
            Write-Host ""
            Write-Host "  Initiating authentication..." -ForegroundColor Yellow
            Write-Host "  (The device code will appear below)" -ForegroundColor Yellow
            Write-Host ""

            # Connect and display output
            try {
                if ($tenantName) {
                    Connect-PnPOnline -Url $connectionUrl -ClientId $body.clientId -Tenant $tenantName -DeviceLogin *>&1 | Out-Host
                } else {
                    Connect-PnPOnline -Url $connectionUrl -ClientId $body.clientId -DeviceLogin *>&1 | Out-Host
                }
            }
            finally {
                [Console]::Out.Flush()
            }

            Write-Host ""
            Write-Host "========================================" -ForegroundColor Green
            Write-Host "  AUTHENTICATION COMPLETED" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Green
            Write-Host ""
        }
        else {
            # Host mode: use interactive browser popup (doesn't need -Tenant parameter)
            # Interactive mode can auto-switch to admin, so use the original URL
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

        # Test user capabilities
        $capabilities = $null
        try {
            Write-ActivityLog "Testing user capabilities..." -Level "Information"
            $capabilities = Test-UserCapabilities
        }
        catch {
            Write-ActivityLog "Capability check failed: $($_.Exception.Message)" -Level "Warning"
            # Return safe defaults if capability check fails
            # This ensures connection succeeds even if capability testing fails
            $capabilities = @{
                CanEnumerateSites = $false
                CanReadUsers = $false
                CanAccessStorageData = $false
                CanReadExternalUsers = $false
                CheckedAt = (Get-Date).ToString("o")
                Error = "Capability check failed"
            }
        }

        Send-JsonResponse -Response $Response -Data @{
            success = $true
            message = "Connected to SharePoint Online"
            siteTitle = if ($web) { $web.Title } else { "SharePoint Site" }
            siteUrl = if ($web) { $web.Url } else { $body.tenantUrl }
            user = $currentUser
            capabilities = $capabilities
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

        $script:ServerState.OperationLog.Clear()
        $script:ServerState.OperationRunning = $true
        $script:ServerState.OperationComplete = $false

        Write-ConsoleOutput "Starting Demo Mode..."
        Write-ConsoleOutput "Generating simulated SharePoint data..."

        # Delegate to DemoDataGenerator
        New-DemoData

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

    if ($script:DemoMode) {
        Send-JsonResponse -Response $Response -Data @{ success = $true; message = "Sites already loaded in demo mode" }
        return
    }

    if ($script:ServerState.OperationRunning) {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = "Another operation is already running" } -StatusCode 409
        return
    }

    # Launch in background so the server stays responsive for progress polling
    $script:ServerState.OperationLog.Clear()
    $script:ServerState.OperationError = $null
    [void]$script:ServerState.OperationLog.Add("Fetching sites...")

    Start-BackgroundOperation -ScriptBlock {
        Get-RealSites-DataDriven
        [void]$SharedState.OperationLog.Add("Sites loaded successfully")
    }

    Send-JsonResponse -Response $Response -Data @{ success = $true; started = $true; message = "Site retrieval started" }
}

# ---- Prepare Analysis (check if re-auth needed) ----

function Handle-PostPrepareAnalysis {
    param($Request, $Response)

    if (-not $script:SPOConnected -and -not $script:DemoMode) {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = "Not connected" } -StatusCode 400
        return
    }

    $body = Read-RequestBody -Request $Request
    $siteUrl = if ($body -and $body.siteUrl) { $body.siteUrl } else { "" }

    if ($script:DemoMode) {
        Send-JsonResponse -Response $Response -Data @{
            needsAuth = $false
            ready = $true
        }
        return
    }

    # Check if re-authentication is needed (container mode + different site)
    $needsReauth = $false
    if ($env:SPO_HEADLESS -and $siteUrl) {
        try {
            $currentUrl = (Get-PnPConnection -ErrorAction SilentlyContinue).Url
            $currentUrlNorm = if ($currentUrl) { $currentUrl.TrimEnd('/') } else { "" }
            $targetUrlNorm = $siteUrl.TrimEnd('/')
            $needsReauth = ($currentUrlNorm -ne $targetUrlNorm)
        }
        catch {
            $needsReauth = $true
        }
    }

    if ($needsReauth) {
        Send-JsonResponse -Response $Response -Data @{
            success = $true
            needsAuth = $true
            currentSite = $currentUrl
            targetSite = $siteUrl
            message = "Re-authentication required for different site"
        }
    }
    else {
        Send-JsonResponse -Response $Response -Data @{
            success = $true
            needsAuth = $false
            ready = $true
        }
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

    if ($script:DemoMode) {
        $script:ServerState.OperationRunning = $false
        $script:ServerState.OperationComplete = $true
        Send-JsonResponse -Response $Response -Data @{ success = $true; message = "Permissions available from demo data" }
        return
    }

    if ($script:ServerState.OperationRunning) {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = "Another operation is already running" } -StatusCode 409
        return
    }

    # Prepare for operation
    $script:ServerState.OperationLog.Clear()
    $script:ServerState.OperationError = $null
    $script:ServerState.OperationSiteUrl = $siteUrl
    $script:ServerState.OperationRunning = $true
    $script:ServerState.OperationComplete = $false

    # Check if re-authentication is needed (container mode, different site)
    $needsReauth = $false
    if ($env:SPO_HEADLESS -and $siteUrl) {
        try {
            $currentUrl = (Get-PnPConnection -ErrorAction SilentlyContinue).Url
            $currentUrlNorm = if ($currentUrl) { $currentUrl.TrimEnd('/') } else { "" }
            $targetUrlNorm = $siteUrl.TrimEnd('/')
            $needsReauth = ($currentUrlNorm -ne $targetUrlNorm)
        }
        catch {
            $needsReauth = $true
        }
    }

    # If re-auth needed, do it before returning response
    if ($needsReauth) {
        [void]$script:ServerState.OperationLog.Add("Authenticating to target site...")

        try {
            Write-Host ""
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "  SITE-SPECIFIC AUTHENTICATION" -ForegroundColor Cyan
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host "  Target site: $siteUrl" -ForegroundColor White
            Write-Host "  Initiating device code authentication..." -ForegroundColor Yellow
            Write-Host ""

            $clientId = Get-AppSetting -SettingName "SharePoint.ClientId"
            $tenantName = ""
            if ($siteUrl -match '//([^-\.]+)') {
                $tenantName = "$($matches[1]).onmicrosoft.com"
            }

            if ($tenantName) {
                Connect-PnPOnline -Url $siteUrl -ClientId $clientId -Tenant $tenantName -DeviceLogin *>&1 | Out-Host
            } else {
                Connect-PnPOnline -Url $siteUrl -ClientId $clientId -DeviceLogin *>&1 | Out-Host
            }
            [Console]::Out.Flush()

            Write-Host ""
            Write-Host "  ✓ Authenticated to target site" -ForegroundColor Green
            Write-Host "========================================" -ForegroundColor Cyan
            Write-Host ""

            [void]$script:ServerState.OperationLog.Add("✅ Authentication successful!")
        }
        catch {
            Write-ActivityLog "Authentication failed: $($_.Exception.Message)" -Level "Error"
            [void]$script:ServerState.OperationLog.Add("❌ Authentication failed: $($_.Exception.Message)")
            $script:ServerState.OperationRunning = $false
            $script:ServerState.OperationComplete = $true
            $script:ServerState.OperationError = "Authentication failed"
            Send-JsonResponse -Response $Response -Data @{ success = $false; message = "Authentication failed: $($_.Exception.Message)" }
            return
        }
    }

    # Return response - authentication is done if it was needed
    Send-JsonResponse -Response $Response -Data @{ success = $true; started = $true; message = "Permissions analysis started" }

    # Start the actual analysis background job
    [void]$script:ServerState.OperationLog.Add("Starting permissions analysis...")
    Write-ActivityLog "Permissions analysis requested for site: '$siteUrl'" -Level "Information"

    Start-BackgroundOperation -ScriptBlock {
        $siteUrl = $SharedState.OperationSiteUrl
        [void]$SharedState.OperationLog.Add("DEBUG: Background job starting with site URL: '$siteUrl'")
        Get-RealPermissions-DataDriven -SiteUrl $siteUrl
        [void]$SharedState.OperationLog.Add("Analysis complete.")
    }
}

# ---- Progress ----

function Handle-GetProgress {
    param($Response)

    $data = @{
        messages = @($script:ServerState.OperationLog.ToArray())
        running  = $script:ServerState.OperationRunning
        complete = $script:ServerState.OperationComplete
    }

    # Include error if the background operation failed
    if ($script:ServerState.OperationError) {
        $data.error = $script:ServerState.OperationError
    }

    # Include enrichment result if available
    if ($script:ServerState.EnrichmentResult) {
        $data.enrichmentResult = $script:ServerState.EnrichmentResult
    }

    Send-JsonResponse -Response $Response -Data $data
}

# ---- Data ----

function Handle-GetData {
    param($Response, [string]$DataType)

    $mappedType = if ($script:DataTypeMap.ContainsKey($DataType.ToLower())) { $script:DataTypeMap[$DataType.ToLower()] } else { $DataType }

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
        $mappedType = if ($script:DataTypeMap.ContainsKey($DataType.ToLower())) { $script:DataTypeMap[$DataType.ToLower()] } else { $DataType }
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

# ---- Graph Enrichment ----

function Handle-PostEnrich {
    param($Response)

    if ($script:DemoMode) {
        # Delegate demo enrichment to DemoDataGenerator
        try {
            $result = Invoke-DemoEnrichment
            Send-JsonResponse -Response $Response -Data @{
                success       = $true
                totalExternal = $result.TotalExternal
                enriched      = $result.Enriched
                failed        = $result.Failed
            }
        }
        catch {
            Send-JsonResponse -Response $Response -Data @{
                success = $false
                message = "Enrichment failed: $($_.Exception.Message)"
            }
        }
        return
    }

    if ($script:ServerState.OperationRunning) {
        Send-JsonResponse -Response $Response -Data @{ success = $false; message = "Another operation is already running" } -StatusCode 409
        return
    }

    # Launch in background so the server stays responsive for progress polling
    $script:ServerState.OperationLog.Clear()
    $script:ServerState.OperationError = $null
    [void]$script:ServerState.OperationLog.Add("Starting external user enrichment...")

    Start-BackgroundOperation -ScriptBlock {
        $result = Invoke-ExternalUserEnrichment
        [void]$SharedState.OperationLog.Add("Enrichment complete: $($result.Enriched) of $($result.TotalExternal) enriched")
        # Store result for the progress endpoint to pick up
        $SharedState.EnrichmentResult = @{
            TotalExternal = $result.TotalExternal
            Enriched      = $result.Enriched
            Failed        = $result.Failed
        }
    }

    Send-JsonResponse -Response $Response -Data @{ success = $true; started = $true; message = "Enrichment started" }
}

function Handle-GetEnrichment {
    param($Response)

    try {
        $summary = Get-EnrichmentSummary

        Send-JsonResponse -Response $Response -Data @{
            totalExternal    = $summary.TotalExternal
            enrichedCount    = $summary.EnrichedCount
            disabledAccounts = $summary.DisabledAccounts
            guestUsers       = $summary.GuestUsers
            staleAccounts    = $summary.StaleAccounts
        }
    }
    catch {
        Send-JsonResponse -Response $Response -Data @{
            totalExternal = 0
            enrichedCount = 0
            error         = $_.Exception.Message
        }
    }
}

# ---- Risk Assessment ----

function Handle-GetRisk {
    param($Response)

    try {
        $assessment = Get-RiskAssessment

        Send-JsonResponse -Response $Response -Data @{
            overallScore  = $assessment.OverallScore
            riskLevel     = $assessment.RiskLevel
            totalFindings = $assessment.TotalFindings
            criticalCount = $assessment.CriticalCount
            highCount     = $assessment.HighCount
            mediumCount   = $assessment.MediumCount
            lowCount      = $assessment.LowCount
            findings      = @($assessment.Findings)
        }
    }
    catch {
        Send-JsonResponse -Response $Response -Data @{
            overallScore  = 0
            riskLevel     = "Unknown"
            totalFindings = 0
            findings      = @()
            error         = $_.Exception.Message
        }
    }
}

# ---- Audit ----

function Handle-GetAudit {
    param($Response)

    $session = Get-AuditSession
    if (-not $session) {
        Send-JsonResponse -Response $Response -Data @{
            hasSession = $false
            message = "No audit session available"
        }
        return
    }

    Send-JsonResponse -Response $Response -Data @{
        hasSession    = $true
        sessionId     = $session.SessionId
        operationType = $session.OperationType
        status        = $session.Status
        startTime     = $session.StartTimestamp
        endTime       = $session.EndTimestamp
        duration      = $session.Duration
        scanScope     = $session.ScanScope
        userPrincipal = $session.UserPrincipal
        errorCount    = $session.ErrorCount
        eventCount    = $session.Events.Count
        outputFiles   = @($session.OutputFiles)
        metrics       = $session.Metrics
    }
}

# ---- Export ----

function Handle-PostExport {
    param($Request, $Response, [string]$ExportType)

    try {
        $mappedType = if ($script:DataTypeMap.ContainsKey($ExportType.ToLower())) { $script:DataTypeMap[$ExportType.ToLower()] } else { $ExportType }
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

function Handle-PostBuildPermissionsMatrix {
    <#
    .SYNOPSIS
    Builds a permissions matrix for a SharePoint site
    #>
    param(
        [System.Net.HttpListenerRequest]$Request,
        [System.Net.HttpListenerResponse]$Response
    )

    try {
        $body = Read-RequestBody -Request $Request
        $siteUrl = $body.siteUrl
        $scanType = $body.scanType

        if (-not $siteUrl) {
            Send-JsonResponse -Response $Response -Data @{ success = $false; message = "Site URL required" } -StatusCode 400
            return
        }

        Write-ActivityLog "Building permissions matrix for site: $siteUrl (type: $scanType)"

        # Verify we're connected
        if (-not $script:SPOConnected -and -not $script:DemoMode) {
            Send-JsonResponse -Response $Response -Data @{
                success = $false
                message = "Not connected to SharePoint. Connect first or use Demo Mode."
            } -StatusCode 400
            return
        }

        # Demo mode - delegate to DemoDataGenerator
        if ($script:DemoMode) {
            Write-ActivityLog "Generating demo permissions matrix data"
            $matrix = Get-DemoPermissionsMatrix -SiteUrl $siteUrl -ScanType $scanType
            Send-JsonResponse -Response $Response -Data @{
                success = $true
                data = $matrix
            }
            return
        }

        # Live mode - connect to the specific site
        try {
            $currentConnection = Get-PnPConnection -ErrorAction SilentlyContinue
            if ($currentConnection) {
                $accessToken = Get-PnPAccessToken
                Connect-PnPOnline -Url $siteUrl -AccessToken $accessToken -WarningAction SilentlyContinue
            } else {
                throw "No active PnP connection"
            }
        } catch {
            Send-JsonResponse -Response $Response -Data @{
                success = $false
                message = "Failed to connect to site: $($_.Exception.Message)"
            } -StatusCode 500
            return
        }

        # Build the permissions matrix
        try {
            $matrix = Get-SitePermissionsMatrix -SiteUrl $siteUrl -ScanType $scanType

            Send-JsonResponse -Response $Response -Data @{
                success = $true
                data = $matrix
            }
        } catch {
            Write-ActivityLog "Matrix build error: $($_.Exception.Message)" -Level "Error"
            Send-JsonResponse -Response $Response -Data @{
                success = $false
                message = $_.Exception.Message
            } -StatusCode 500
        }

    } catch {
        Write-ActivityLog "Matrix build error: $($_.Exception.Message)" -Level "Error"
        Send-JsonResponse -Response $Response -Data @{
            success = $false
            message = $_.Exception.Message
        } -StatusCode 500
    }
}
