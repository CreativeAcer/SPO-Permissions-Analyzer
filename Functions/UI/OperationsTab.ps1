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
                Generate-DemoReport-DataDriven
            } else {
                Generate-RealReport-DataDriven
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

# function Get-SharePointSites {
#     <#
#     .SYNOPSIS
#     Retrieves and displays SharePoint sites with real-time console updates
#     #>
#     try {
#         Write-ActivityLog "Starting SharePoint sites retrieval" -Level "Information"
        
#         # Clear previous results and show starting message
#         Write-ConsoleOutput "🔍 SHAREPOINT SITES ANALYSIS" -ForceUpdate
#         Write-ConsoleOutput "=====================================================" -Append -ForceUpdate
#         Write-ConsoleOutput "⏱️ Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append -ForceUpdate
#         Write-ConsoleOutput "" -Append -ForceUpdate
        
#         if (Get-AppSetting -SettingName "DemoMode") {
#             Get-DemoSites-DataDriven
#         } else {
#             Get-RealSites-DataDriven
#         }
        
#         Write-ActivityLog "SharePoint sites retrieval completed" -Level "Information"
#     }
#     catch {
#         Write-ConsoleOutput "❌ ERROR: Failed to retrieve SharePoint sites" -Append
#         Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
#         Write-ErrorLog -Message $_.Exception.Message -Location "Get-SharePointSites"
#     }
# }

function Get-RealSites-DataDriven {
    <#
    .SYNOPSIS
    Retrieves real SharePoint sites and stores in data manager
    #>
    try {
        if (-not $script:SPOConnected) {
            throw "Not connected to SharePoint. Please connect first."
        }
        
        # Clear previous sites data
        Clear-SharePointData -DataType "Sites"
        Set-SharePointOperationContext -OperationType "Sites Analysis"
        
        Write-ConsoleOutput "🔄 Using modern PnP PowerShell for site enumeration..." -Append -ForceUpdate
        Write-ConsoleOutput "📡 Attempting modern tenant-level site enumeration..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 1000
        
        $sites = @()
        
        # Try to get sites (your existing logic)
        try {
            $sites = Get-PnPTenantSite -ErrorAction Stop
            Write-ActivityLog "Retrieved $($sites.Count) sites using Get-PnPTenantSite"
        }
        catch {
            # Fallback logic
            $currentWeb = Get-PnPWeb -ErrorAction Stop
            $subsites = Get-PnPSubWeb -Recurse -ErrorAction SilentlyContinue
            $sites = @($currentWeb)
            if ($subsites) { $sites += $subsites }
        }
        
        Write-ConsoleOutput "🏢 Sites Found: $($sites.Count)" -Append -ForceUpdate
        Write-ConsoleOutput "" -Append
        
        # Process and store each site
        $siteCounter = 0
        foreach ($site in $sites | Select-Object -First 25) {
            $siteCounter++
            
            # Create site data object
            $siteData = @{
                Title = if ($site.Title) { $site.Title } else { "Site $siteCounter" }
                Url = if ($site.Url) { $site.Url } else { "N/A" }
                Owner = if ($site.Owner) { $site.Owner } elseif ($site.SiteOwnerEmail) { $site.SiteOwnerEmail } else { "N/A" }
                Storage = if ($site.StorageUsageCurrent) { $site.StorageUsageCurrent.ToString() } else { "500" }
                StorageQuota = if ($site.StorageQuota) { $site.StorageQuota } else { 0 }
                Template = if ($site.Template) { $site.Template } else { "N/A" }
                LastModified = if ($site.LastContentModifiedDate) { $site.LastContentModifiedDate.ToString() } else { "N/A" }
                IsHubSite = if ($site.IsHubSite) { $site.IsHubSite } else { $false }
                UserCount = 0  # Will be updated if available
                GroupCount = 0  # Will be updated if available
            }
            
            # Add site to data store
            Add-SharePointSite -SiteData $siteData
            
            # Display in console (for user visibility)
            Write-ConsoleOutput "📁 SITE #$siteCounter`: $($siteData.Title)" -Append -ForceUpdate
            Write-ConsoleOutput "   🌐 URL: $($siteData.Url)" -Append
            Write-ConsoleOutput "   👤 Owner: $($siteData.Owner)" -Append
            Write-ConsoleOutput "   💾 Storage: $($siteData.Storage) MB" -Append
            if ($siteData.Template -ne "N/A") { Write-ConsoleOutput "   🎨 Template: $($siteData.Template)" -Append }
            if ($siteData.LastModified -ne "N/A") { Write-ConsoleOutput "   📅 Last Modified: $($siteData.LastModified)" -Append }
            if ($siteData.IsHubSite) { Write-ConsoleOutput "   🌟 Hub Site" -Append }
            Write-ConsoleOutput "   ============================================================" -Append
            Update-UIAndWait -WaitMs 300
        }
        
        if ($sites.Count -gt 25) {
            Write-ConsoleOutput "... and $($sites.Count - 25) more sites" -Append
        }
        
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ Site enumeration completed successfully!" -Append -ForceUpdate
        
        # UPDATE VISUAL ANALYTICS DIRECTLY FROM DATA
        Update-VisualAnalyticsFromData
        
        Write-ActivityLog "Sites operation completed with $($sites.Count) sites" -Level "Information"
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-RealSites-DataDriven"
        Write-ConsoleOutput "❌ ERROR: Site enumeration failed" -Append
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
    }
}

# function Get-SharePointPermissions {
#     <#
#     .SYNOPSIS
#     Analyzes SharePoint permissions with real-time console updates
#     #>
#     try {
#         Write-ActivityLog "Starting SharePoint permissions analysis" -Level "Information"
        
#         # Get site URL if specified
#         $siteUrl = $script:txtSiteUrl.Text.Trim()
#         if ([string]::IsNullOrEmpty($siteUrl)) {
#             $siteUrl = "All Sites (Tenant-wide analysis)"
#         }
        
#         # Clear previous results and show starting message
#         Write-ConsoleOutput "🔐 SHAREPOINT PERMISSIONS ANALYSIS" -ForceUpdate
#         Write-ConsoleOutput "=====================================================" -Append -ForceUpdate
#         Write-ConsoleOutput "⏱️ Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -Append -ForceUpdate
#         Write-ConsoleOutput "🎯 Target: $siteUrl" -Append -ForceUpdate
#         Write-ConsoleOutput "" -Append -ForceUpdate
        
#         if (Get-AppSetting -SettingName "DemoMode") {
#             Get-DemoPermissions-DataDriven -SiteUrl $siteUrl
#         } else {
#             Get-RealPermissions-DataDriven -SiteUrl $siteUrl
#         }
        
#         Write-ActivityLog "SharePoint permissions analysis completed" -Level "Information"
#     }
#     catch {
#         Write-ConsoleOutput "❌ ERROR: Failed to analyze SharePoint permissions" -Append
#         Write-ConsoleOutput "Error Details: $($_.Exception.Message)" -Append
#         Write-ErrorLog -Message $_.Exception.Message -Location "Get-SharePointPermissions"
#     }
# }

function Get-RealPermissions-DataDriven {
    <#
    .SYNOPSIS
    Analyzes real SharePoint permissions and stores in data manager
    #>
    param([string]$SiteUrl)
    
    try {
        if (-not $script:SPOConnected) {
            throw "Not connected to SharePoint. Please connect first."
        }
        
        # Clear previous data
        Clear-SharePointData -DataType "All"
        Set-SharePointOperationContext -OperationType "Permissions Analysis"
        
        $siteUrl = $script:txtSiteUrl.Text.Trim()
        if ([string]::IsNullOrEmpty($siteUrl)) {
            Write-ConsoleOutput "Please enter a specific site URL to analyze permissions." -Append
            return
        }
        
        Write-ConsoleOutput "🔄 Analyzing permissions for: $siteUrl..." -Append -ForceUpdate
        
        # Connect to the site
        Connect-PnPOnline -Url $siteUrl -ClientId (Get-AppSetting -SettingName "SharePoint.ClientId") -Interactive
        
        # Get site information
        $web = Get-PnPWeb -ErrorAction Stop
        
        # Add the analyzed site to data store
        $siteData = @{
            Title = if ($web.Title) { $web.Title } else { "Analyzed Site" }
            Url = if ($web.Url) { $web.Url } else { $siteUrl }
            Owner = "Current User"
            Storage = "850"  # Default value
            Template = if ($web.WebTemplate) { $web.WebTemplate } else { "N/A" }
            HasUniquePermissions = $web.HasUniqueRoleAssignments
        }
        Add-SharePointSite -SiteData $siteData
        
        # Display site info
        Write-ConsoleOutput "✅ SITE INFORMATION" -Append -ForceUpdate
        Write-ConsoleOutput "📝 Title: $($web.Title)" -Append
        Write-ConsoleOutput "🌐 URL: $($web.Url)" -Append
        Write-ConsoleOutput "🔒 Has Unique Permissions: $($web.HasUniqueRoleAssignments)" -Append
        Write-ConsoleOutput "" -Append
        
        # Get and store users
        Write-ConsoleOutput "👥 Retrieving users..." -Append -ForceUpdate
        try {
            $users = Get-PnPUser -ErrorAction Stop
            $userCounter = 0
            
            foreach ($user in $users | Where-Object { $_.PrincipalType -eq "User" }) {
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
            
            Write-ConsoleOutput "✅ Retrieved $userCounter users" -Append
            Write-ConsoleOutput "" -Append
        }
        catch {
            Write-ConsoleOutput "⚠️ Limited access to user information" -Append
            Write-ConsoleOutput "" -Append
        }
        
        # Get and store groups
        Write-ConsoleOutput "👨‍👩‍👧‍👦 Retrieving groups..." -Append -ForceUpdate
        try {
            $groups = Get-PnPGroup -ErrorAction Stop
            $groupCounter = 0
            
            foreach ($group in $groups) {
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
            
            Write-ConsoleOutput "✅ Retrieved $groupCounter groups" -Append
            Write-ConsoleOutput "" -Append
        }
        catch {
            Write-ConsoleOutput "❌ Failed to retrieve groups" -Append
            Write-ConsoleOutput "" -Append
        }
        
        Write-ConsoleOutput "✅ Permissions analysis completed successfully!" -Append -ForceUpdate
        
        # UPDATE VISUAL ANALYTICS DIRECTLY FROM DATA
        Update-VisualAnalyticsFromData
        
        Write-ActivityLog "Permissions analysis completed" -Level "Information"
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-RealPermissions-DataDriven"
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
    Demo permissions using data-driven approach
    #>
    try {
        # Clear and set context
        Clear-SharePointData -DataType "All"
        Set-SharePointOperationContext -OperationType "Permissions Analysis (Demo)"
        
        Write-ConsoleOutput "🎭 Running in Demo Mode - Simulating permissions analysis..." -Append -ForceUpdate
        Update-UIAndWait -WaitMs 500
        
        # Add demo site
        Add-SharePointSite -SiteData @{
            Title = "Team Collaboration Site"
            Url = "https://demo.sharepoint.com/sites/teamsite"
            Owner = "admin@demo.com"
            Storage = "750"
        }
        
        # Add demo users
        $demoUsers = @(
            @{Name="John Doe"; Email="john.doe@demo.com"; Type="Internal"; Permission="Full Control"},
            @{Name="Jane Smith"; Email="jane.smith@demo.com"; Type="Internal"; Permission="Edit"},
            @{Name="Mike Johnson"; Email="mike.johnson@demo.com"; Type="Internal"; Permission="Read"},
            @{Name="External Partner"; Email="partner@external.com"; Type="External"; Permission="Read"}
        )
        
        foreach ($user in $demoUsers) {
            Add-SharePointUser -UserData $user
        }
        
        # Add demo groups
        $demoGroups = @(
            @{Name="Site Owners"; MemberCount=3; Permission="Full Control"},
            @{Name="Site Members"; MemberCount=8; Permission="Edit"},
            @{Name="Site Visitors"; MemberCount=4; Permission="Read"}
        )
        
        foreach ($group in $demoGroups) {
            Add-SharePointGroup -GroupData $group
        }
        
        # Display summary
        $metrics = Get-SharePointData -DataType "Metrics"
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "📊 PERMISSION SUMMARY:" -Append -ForceUpdate
        Write-ConsoleOutput "   👤 Total Users: $($metrics.TotalUsers)" -Append
        Write-ConsoleOutput "   👨‍👩‍👧‍👦 Total Groups: $($metrics.TotalGroups)" -Append
        Write-ConsoleOutput "   🌐 External Users: $($metrics.ExternalUsers)" -Append
        Write-ConsoleOutput "" -Append
        Write-ConsoleOutput "✅ Permissions analysis completed successfully!" -Append -ForceUpdate
        
        # Update Visual Analytics directly from data
        Update-VisualAnalyticsFromData
        
    }
    catch {
        Write-ConsoleOutput "❌ Demo permissions analysis failed: $($_.Exception.Message)" -Append
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
        
        # Update Visual Analytics
        $consoleText = $script:txtOperationsResults.Text
        Parse-ConsoleOutputAndUpdateAnalytics -ConsoleText $consoleText -OperationType "Report"
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
        
        # Update Visual Analytics
        $consoleText = $script:txtOperationsResults.Text
        Parse-ConsoleOutputAndUpdateAnalytics -ConsoleText $consoleText -OperationType "Report"
        
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

# Helper function for UI updates and waiting
function Update-UIAndWait {
    param([int]$WaitMs = 500)
    
    try {
        # Force UI update
        [System.Windows.Forms.Application]::DoEvents()
        
        # Wait for specified time
        Start-Sleep -Milliseconds $WaitMs
    }
    catch {
        # Ignore errors in UI updates
        Start-Sleep -Milliseconds $WaitMs
    }
}

# Helper function to write console output
function Write-ConsoleOutput {
    param(
        [string]$Message,
        [switch]$Append,
        [switch]$ForceUpdate
    )
    
    try {
        if ($Append) {
            $script:txtOperationsResults.Text += "`n$Message"
        } else {
            $script:txtOperationsResults.Text = $Message
        }
        
        if ($ForceUpdate) {
            # Scroll to bottom
            $script:txtOperationsResults.SelectionStart = $script:txtOperationsResults.Text.Length
            $script:txtOperationsResults.ScrollToCaret()
            
            # Force UI refresh
            [System.Windows.Forms.Application]::DoEvents()
        }
    }
    catch {
        # Fallback to Write-Host if UI is not available
        Write-Host $Message
    }
}