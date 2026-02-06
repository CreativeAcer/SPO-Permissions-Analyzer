function Initialize-VisualAnalyticsTab {
    <#
    .SYNOPSIS
    Initializes the Visual Analytics tab with charts and data displays
    #>
    try {
        Write-ActivityLog "Initializing Visual Analytics tab" -Level "Information"

        $script:btnRefreshAnalytics.Add_Click({
            Update-VisualAnalyticsFromData
        })

        # Make metric tiles clickable
        Make-MetricTilesClickable
        
        # Initialize with empty state
        Reset-VisualAnalytics
        
        Write-ActivityLog "Visual Analytics tab initialized successfully" -Level "Information"
    }
    catch {
        Write-ActivityLog "Failed to initialize Visual Analytics tab: $($_.Exception.Message)" -Level "Error"
        throw
    }
}

function Make-MetricTilesClickable {
    <#
    .SYNOPSIS
    Makes the metric tiles clickable to open deep dive windows
    #>
    try {
        # Find the metric card borders (parent containers of the metrics)
        $gridMetricsCards = $script:MainWindow.FindName("gridMetricsCards")

        if ($gridMetricsCards) {
            # Total Sites Card (Column 0)
            $sitesCard = $gridMetricsCards.Children | Where-Object {
                [System.Windows.Controls.Grid]::GetColumn($_) -eq 0
            }
            if ($sitesCard) {
                $sitesCard.Cursor = [System.Windows.Input.Cursors]::Hand
                $sitesCard.Add_MouseLeftButtonUp({
                    Open-SitesDeepDive
                })
                $sitesCard.Add_MouseEnter({
                    $this.Opacity = 0.8
                })
                $sitesCard.Add_MouseLeave({
                    $this.Opacity = 1.0
                })
            }

            # Total Users Card (Column 1)
            $usersCard = $gridMetricsCards.Children | Where-Object {
                [System.Windows.Controls.Grid]::GetColumn($_) -eq 1
            }
            if ($usersCard) {
                $usersCard.Cursor = [System.Windows.Input.Cursors]::Hand
                $usersCard.Add_MouseLeftButtonUp({
                    Open-UsersDeepDive
                })
                $usersCard.Add_MouseEnter({
                    $this.Opacity = 0.8
                })
                $usersCard.Add_MouseLeave({
                    $this.Opacity = 1.0
                })
            }

            # Total Groups Card (Column 2)
            $groupsCard = $gridMetricsCards.Children | Where-Object {
                [System.Windows.Controls.Grid]::GetColumn($_) -eq 2
            }
            if ($groupsCard) {
                $groupsCard.Cursor = [System.Windows.Input.Cursors]::Hand
                $groupsCard.Add_MouseLeftButtonUp({
                    Open-GroupsDeepDive
                })
                $groupsCard.Add_MouseEnter({
                    $this.Opacity = 0.8
                })
                $groupsCard.Add_MouseLeave({
                    $this.Opacity = 1.0
                })
            }

            # External Users Card (Column 3)
            $externalCard = $gridMetricsCards.Children | Where-Object {
                [System.Windows.Controls.Grid]::GetColumn($_) -eq 3
            }
            if ($externalCard) {
                $externalCard.Cursor = [System.Windows.Input.Cursors]::Hand
                $externalCard.Add_MouseLeftButtonUp({
                    Open-ExternalUsersDeepDive
                })
                $externalCard.Add_MouseEnter({
                    $this.Opacity = 0.8
                })
                $externalCard.Add_MouseLeave({
                    $this.Opacity = 1.0
                })
            }

            Write-ActivityLog "Metric tiles made clickable" -Level "Information"
        }
        else {
            Write-ActivityLog "Could not find metrics cards grid" -Level "Warning"
        }

        # Security Analysis Cards (P3 features)
        $gridSecurityCards = $script:MainWindow.FindName("gridSecurityCards")
        if ($gridSecurityCards) {
            # Role Assignments Card (Column 0)
            $roleCard = $gridSecurityCards.Children | Where-Object {
                [System.Windows.Controls.Grid]::GetColumn($_) -eq 0
            }
            if ($roleCard) {
                $roleCard.Cursor = [System.Windows.Input.Cursors]::Hand
                $roleCard.Add_MouseLeftButtonUp({ Open-PermissionsDeepDive })
                $roleCard.Add_MouseEnter({ $this.Opacity = 0.8 })
                $roleCard.Add_MouseLeave({ $this.Opacity = 1.0 })
            }

            # Inheritance Breaks Card (Column 1)
            $inheritanceCard = $gridSecurityCards.Children | Where-Object {
                [System.Windows.Controls.Grid]::GetColumn($_) -eq 1
            }
            if ($inheritanceCard) {
                $inheritanceCard.Cursor = [System.Windows.Input.Cursors]::Hand
                $inheritanceCard.Add_MouseLeftButtonUp({ Open-InheritanceDeepDive })
                $inheritanceCard.Add_MouseEnter({ $this.Opacity = 0.8 })
                $inheritanceCard.Add_MouseLeave({ $this.Opacity = 1.0 })
            }

            # Sharing Links Card (Column 2)
            $sharingCard = $gridSecurityCards.Children | Where-Object {
                [System.Windows.Controls.Grid]::GetColumn($_) -eq 2
            }
            if ($sharingCard) {
                $sharingCard.Cursor = [System.Windows.Input.Cursors]::Hand
                $sharingCard.Add_MouseLeftButtonUp({ Open-SharingLinksDeepDive })
                $sharingCard.Add_MouseEnter({ $this.Opacity = 0.8 })
                $sharingCard.Add_MouseLeave({ $this.Opacity = 1.0 })
            }

            Write-ActivityLog "Security analysis tiles made clickable" -Level "Information"
        }

    }
    catch {
        Write-ActivityLog "Error making metric tiles clickable: $($_.Exception.Message)" -Level "Warning"
    }
}

function Reset-VisualAnalytics {
    <#
    .SYNOPSIS
    Resets Visual Analytics to initial empty state
    #>
    try {
        # Reset metrics cards
        $script:txtTotalSites.Text = "0"
        $script:txtTotalUsers.Text = "0"
        $script:txtTotalGroups.Text = "0"
        $script:txtExternalUsers.Text = "0"

        # Reset security analysis cards
        if ($script:txtRoleAssignments) { $script:txtRoleAssignments.Text = "0" }
        if ($script:txtInheritanceBreaks) { $script:txtInheritanceBreaks.Text = "0" }
        if ($script:txtSharingLinks) { $script:txtSharingLinks.Text = "0" }

        # Clear data displays
        $script:dgSites.ItemsSource = $null
        $script:lstPermissionAlerts.ItemsSource = $null

        # Clear and reset charts
        Reset-StorageChart
        Reset-PermissionChart

        # Set appropriate subtitle
        $script:txtAnalyticsSubtitle.Text = "Run an analysis operation to view visual insights"

        Write-ActivityLog "Visual Analytics reset to initial state" -Level "Information"
    }
    catch {
        Write-ActivityLog "Error resetting Visual Analytics: $($_.Exception.Message)" -Level "Warning"
    }
}

function Update-Charts {
    <#
    .SYNOPSIS
    Updates both storage and permission charts
    #>
    param(
        [array]$SitesData
    )
    
    try {
        Update-StorageChart -SitesData $SitesData
        Update-PermissionChart
        
        Write-ActivityLog "Charts updated successfully" -Level "Information"
    }
    catch {
        Write-ActivityLog "Error updating charts: $($_.Exception.Message)" -Level "Warning"
    }
}

function Update-StorageChart {
    <#
    .SYNOPSIS
    Updates the storage usage chart with site data
    #>
    param(
        [array]$SitesData
    )
    
    try {
        $script:canvasStorageChart.Children.Clear()
        
        if (-not $SitesData -or $SitesData.Count -eq 0) {
            Add-ChartPlaceholder -Canvas $script:canvasStorageChart -Text "No site data available for chart"
            return
        }
        
        # Filter sites that have storage data
        $sitesWithStorage = @()
        foreach ($site in $SitesData) {
            if ($site["Storage"] -and $site["Storage"] -ne "N/A") {
                $sitesWithStorage += $site
            }
        }
        
        if ($sitesWithStorage.Count -eq 0) {
            Add-ChartPlaceholder -Canvas $script:canvasStorageChart -Text "No storage data available"
            return
        }
        
        # Create simple bar chart
        $maxStorage = 0
        foreach ($site in $sitesWithStorage) {
            $storageValue = [int]($site["Storage"] -replace " MB", "")
            if ($storageValue -gt $maxStorage) { $maxStorage = $storageValue }
        }
        if ($maxStorage -eq 0) { $maxStorage = 1000 } # Prevent division by zero
        
        $barWidth = 30
        $barSpacing = 40
        $chartHeight = 150
        $startX = 10
        
        for ($i = 0; $i -lt [Math]::Min($sitesWithStorage.Count, 5); $i++) {
            $site = $sitesWithStorage[$i]
            $storageValue = [int]($site["Storage"] -replace " MB", "")
            $barHeight = ($storageValue / $maxStorage) * $chartHeight
            
            # Create storage bar
            $bar = New-Object System.Windows.Shapes.Rectangle
            $bar.Width = $barWidth
            $bar.Height = $barHeight
            $bar.Fill = if ($site["UsageColor"]) { $site["UsageColor"] } else { "#007ACC" }
            
            # Position bar
            [System.Windows.Controls.Canvas]::SetLeft($bar, $startX + ($i * $barSpacing))
            [System.Windows.Controls.Canvas]::SetTop($bar, $chartHeight - $barHeight + 10)
            
            $script:canvasStorageChart.Children.Add($bar)
            
            # Add site name label
            $label = New-Object System.Windows.Controls.TextBlock
            $siteTitle = if ($site["Title"]) { $site["Title"] } else { "Site $($i+1)" }
            $label.Text = $siteTitle.Substring(0, [Math]::Min(8, $siteTitle.Length))
            $label.FontSize = 9
            $label.Foreground = "#495057"
            
            [System.Windows.Controls.Canvas]::SetLeft($label, $startX + ($i * $barSpacing) - 5)
            [System.Windows.Controls.Canvas]::SetTop($label, $chartHeight + 15)
            
            $script:canvasStorageChart.Children.Add($label)
            
            # Add storage value label
            $valueLabel = New-Object System.Windows.Controls.TextBlock
            $valueLabel.Text = "$($storageValue)MB"
            $valueLabel.FontSize = 8
            $valueLabel.Foreground = "#666"
            
            [System.Windows.Controls.Canvas]::SetLeft($valueLabel, $startX + ($i * $barSpacing) - 5)
            [System.Windows.Controls.Canvas]::SetTop($valueLabel, $chartHeight - $barHeight - 5)
            
            $script:canvasStorageChart.Children.Add($valueLabel)
        }
        
        Write-ActivityLog "Storage chart updated with $($sitesWithStorage.Count) sites" -Level "Information"
    }
    catch {
        Write-ActivityLog "Error updating storage chart: $($_.Exception.Message)" -Level "Warning"
        Add-ChartPlaceholder -Canvas $script:canvasStorageChart -Text "Error loading storage chart"
    }
}

function Update-PermissionChart {
    <#
    .SYNOPSIS
    Updates the permission distribution chart using real data from the data manager
    #>
    try {
        $script:canvasPermissionChart.Children.Clear()

        # Build permission distribution from actual user and group data
        $permissionCounts = @{}

        $users = Get-SharePointData -DataType "Users"
        foreach ($user in $users) {
            $perm = if ($user["Permission"]) { $user["Permission"] } else { "Unknown" }
            if ($permissionCounts.ContainsKey($perm)) {
                $permissionCounts[$perm]++
            } else {
                $permissionCounts[$perm] = 1
            }
        }

        $groups = Get-SharePointData -DataType "Groups"
        foreach ($group in $groups) {
            $perm = if ($group["Permission"]) { $group["Permission"] } else { "Unknown" }
            $memberCount = if ($group["MemberCount"]) { [int]$group["MemberCount"] } else { 1 }
            if ($permissionCounts.ContainsKey($perm)) {
                $permissionCounts[$perm] += $memberCount
            } else {
                $permissionCounts[$perm] = $memberCount
            }
        }

        if ($permissionCounts.Count -eq 0) {
            Add-ChartPlaceholder -Canvas $script:canvasPermissionChart -Text "No permission data available - Run permission analysis first"
            return
        }

        # Map permission names to colors
        $colorMap = @{
            "Full Control"     = "#DC3545"
            "Edit"             = "#FFC107"
            "Read"             = "#28A745"
            "View Only"        = "#17A2B8"
            "Member"           = "#FFC107"
            "Group Permission" = "#6F42C1"
            "Unknown"          = "#6C757D"
        }

        # Build sorted permissions array
        $permissions = @()
        foreach ($key in $permissionCounts.Keys | Sort-Object) {
            $color = if ($colorMap.ContainsKey($key)) { $colorMap[$key] } else { "#6C757D" }
            $permissions += @{ Name = $key; Count = $permissionCounts[$key]; Color = $color }
        }

        $total = ($permissions | Measure-Object -Property Count -Sum).Sum
        if ($total -eq 0) { $total = 1 }

        $startY = 20
        $barHeight = 25
        $maxBarWidth = 150

        for ($i = 0; $i -lt $permissions.Count; $i++) {
            $perm = $permissions[$i]
            $barWidth = [Math]::Max(($perm.Count / $total) * $maxBarWidth, 5)

            $bar = New-Object System.Windows.Shapes.Rectangle
            $bar.Width = $barWidth
            $bar.Height = $barHeight
            $bar.Fill = $perm.Color

            [System.Windows.Controls.Canvas]::SetLeft($bar, 10)
            [System.Windows.Controls.Canvas]::SetTop($bar, $startY + ($i * ($barHeight + 10)))
            $script:canvasPermissionChart.Children.Add($bar)

            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = "$($perm.Name): $($perm.Count)"
            $label.FontSize = 10
            $label.Foreground = "#495057"

            [System.Windows.Controls.Canvas]::SetLeft($label, $barWidth + 20)
            [System.Windows.Controls.Canvas]::SetTop($label, $startY + ($i * ($barHeight + 10)) + 5)
            $script:canvasPermissionChart.Children.Add($label)
        }

        Write-ActivityLog "Permission chart updated with real data ($($permissions.Count) permission levels)" -Level "Information"
    }
    catch {
        Write-ActivityLog "Error updating permission chart: $($_.Exception.Message)" -Level "Warning"
        Add-ChartPlaceholder -Canvas $script:canvasPermissionChart -Text "Error loading permission chart"
    }
}

function Reset-StorageChart {
    <#
    .SYNOPSIS
    Resets the storage chart to placeholder state
    #>
    try {
        $script:canvasStorageChart.Children.Clear()
        Add-ChartPlaceholder -Canvas $script:canvasStorageChart -Text "No data available - Run site analysis to view storage usage"
    }
    catch {
        Write-ActivityLog "Error resetting storage chart: $($_.Exception.Message)" -Level "Warning"
    }
}

function Reset-PermissionChart {
    <#
    .SYNOPSIS
    Resets the permission chart to placeholder state
    #>
    try {
        $script:canvasPermissionChart.Children.Clear()
        Add-ChartPlaceholder -Canvas $script:canvasPermissionChart -Text "No data available - Run permission analysis to view distribution"
    }
    catch {
        Write-ActivityLog "Error resetting permission chart: $($_.Exception.Message)" -Level "Warning"
    }
}

function Add-ChartPlaceholder {
    <#
    .SYNOPSIS
    Adds placeholder text to empty chart canvas
    #>
    param(
        [System.Windows.Controls.Canvas]$Canvas,
        [string]$Text
    )
    
    try {
        $Canvas.Children.Clear()
        
        $textBlock = New-Object System.Windows.Controls.TextBlock
        $textBlock.Text = $Text
        $textBlock.FontSize = 11
        $textBlock.Foreground = "#6C757D"
        $textBlock.TextWrapping = "Wrap"
        $textBlock.TextAlignment = "Center"
        $textBlock.Width = 180
        
        # Center the text in the canvas
        [System.Windows.Controls.Canvas]::SetLeft($textBlock, 10)
        [System.Windows.Controls.Canvas]::SetTop($textBlock, 80)
        
        $Canvas.Children.Add($textBlock)
    }
    catch {
        Write-ActivityLog "Error adding chart placeholder: $($_.Exception.Message)" -Level "Warning"
    }
}

function Generate-AnalyticsAlerts {
    <#
    .SYNOPSIS
    Generates comprehensive analytics alerts based on parsed data
    #>
    param(
        [int]$SitesCount,
        [int]$UsersCount,
        [int]$ExternalCount,
        [int]$GroupsCount = 0,
        [int]$RecordsProcessed = 0,
        [int]$SecurityFindings = 0
    )
    
    try {
        $alerts = @()
        
        # Security-focused alerts
        if ($ExternalCount -gt 0) {
            $severity = if ($ExternalCount -gt 5) { "#DC3545" } else { "#FFC107" }
            $alerts += [PSCustomObject]@{
                AlertText = "🌐 External Users Detected"
                AlertDetails = "$ExternalCount external users found - security review recommended"
                AlertColor = $severity
            }
        }
        
        if ($SecurityFindings -gt 0) {
            $alerts += [PSCustomObject]@{
                AlertText = "🚨 Security Issues Found"
                AlertDetails = "$SecurityFindings security findings require attention"
                AlertColor = "#DC3545"
            }
        }
        
        # Scale and governance alerts
        if ($UsersCount -gt 50) {
            $alerts += [PSCustomObject]@{
                AlertText = "📈 High User Count"
                AlertDetails = "$UsersCount users detected - consider permission optimization"
                AlertColor = "#FFC107"
            }
        }
        
        if ($SitesCount -gt 10) {
            $alerts += [PSCustomObject]@{
                AlertText = "🏢 Multiple Sites Found"
                AlertDetails = "$SitesCount sites in tenant - governance review recommended"
                AlertColor = "#17A2B8"
            }
        }
        
        # Performance alerts
        if ($RecordsProcessed -gt 1000) {
            $alerts += [PSCustomObject]@{
                AlertText = "📊 Large Dataset Processed"
                AlertDetails = "$RecordsProcessed permission records analyzed successfully"
                AlertColor = "#6F42C1"
            }
        }
        
        # Positive findings
        if ($ExternalCount -eq 0) {
            $alerts += [PSCustomObject]@{
                AlertText = "✅ No External Users"
                AlertDetails = "No external users detected - good security posture"
                AlertColor = "#28A745"
            }
        }
        
        # Completion alert with comprehensive summary
        $alerts += [PSCustomObject]@{
            AlertText = "🎯 Analysis Complete"
            AlertDetails = "Successfully analyzed $SitesCount sites, $UsersCount users, and $GroupsCount groups"
            AlertColor = "#28A745"
        }
        
        # Data freshness alert
        $alerts += [PSCustomObject]@{
            AlertText = "⏰ Data Freshness"
            AlertDetails = "Last updated: $(Get-Date -Format 'HH:mm:ss') - Data is current and actionable"
            AlertColor = "#6C757D"
        }
        
        # Recommendation alert
        if ($SitesCount -gt 0 -or $UsersCount -gt 0) {
            $alerts += [PSCustomObject]@{
                AlertText = "💡 Recommendations Available"
                AlertDetails = "Review console output for detailed recommendations and next steps"
                AlertColor = "#17A2B8"
            }
        }
        
        $script:lstPermissionAlerts.ItemsSource = $alerts
        
        Write-ActivityLog "Generated $($alerts.Count) comprehensive analytics alerts" -Level "Information"
    }
    catch {
        Write-ActivityLog "Error generating analytics alerts: $($_.Exception.Message)" -Level "Warning"
    }
}

function Show-DemoAnalytics {
    <#
    .SYNOPSIS
    Shows demo analytics data when in demo mode
    #>
    try {
        # Update metrics with demo values
        $script:txtTotalSites.Text = "5"
        $script:txtTotalUsers.Text = "47"
        $script:txtTotalGroups.Text = "23"
        $script:txtExternalUsers.Text = "8"
        
        # Create demo sites for data grid
        $demoSites = @(
            [PSCustomObject]@{
                Title = "Team Collaboration Site"
                Url = "https://demo.sharepoint.com/sites/teamsite"
                Owner = "admin@demo.com"
                Storage = "245 MB"
                UsageLevel = "Low"
                UsageColor = "#28A745"
                UserCount = 12
                GroupCount = 4
            },
            [PSCustomObject]@{
                Title = "Project Alpha Workspace"
                Url = "https://demo.sharepoint.com/sites/project-alpha"
                Owner = "project.manager@demo.com"
                Storage = "1024 MB"
                UsageLevel = "High"
                UsageColor = "#DC3545"
                UserCount = 8
                GroupCount = 3
            },
            [PSCustomObject]@{
                Title = "HR Document Center"
                Url = "https://demo.sharepoint.com/sites/hr-documents"
                Owner = "hr.admin@demo.com"
                Storage = "512 MB"
                UsageLevel = "Medium"
                UsageColor = "#FFC107"
                UserCount = 6
                GroupCount = 2
            },
            [PSCustomObject]@{
                Title = "Company Intranet Portal"
                Url = "https://demo.sharepoint.com"
                Owner = "admin@demo.com"
                Storage = "2048 MB"
                UsageLevel = "Critical"
                UsageColor = "#6F42C1"
                UserCount = 47
                GroupCount = 8
            },
            [PSCustomObject]@{
                Title = "Sales Team Hub"
                Url = "https://demo.sharepoint.com/sites/sales"
                Owner = "sales.manager@demo.com"
                Storage = "768 MB"
                UsageLevel = "Medium"
                UsageColor = "#FFC107"
                UserCount = 15
                GroupCount = 5
            }
        )
        
        $script:dgSites.ItemsSource = $demoSites
        
        # Convert demo sites to format expected by chart functions
        $demoSitesForChart = @()
        foreach ($site in $demoSites) {
            $demoSitesForChart += @{
                Title = $site.Title
                Url = $site.Url
                Owner = $site.Owner
                Storage = $site.Storage -replace " MB", ""
                UsageLevel = $site.UsageLevel
                UsageColor = $site.UsageColor
            }
        }
        
        # Update charts with demo data
        Update-StorageChart -SitesData $demoSitesForChart
        Update-PermissionChart
        
        # Generate demo alerts
        Generate-AnalyticsAlerts -SitesCount 5 -UsersCount 47 -ExternalCount 8 -GroupsCount 23 -RecordsProcessed 1247 -SecurityFindings 8
        
        # Update subtitle
        $script:txtAnalyticsSubtitle.Text = "Demo data loaded - showing sample analytics dashboard"
        
        Write-ActivityLog "Demo analytics displayed successfully" -Level "Information"
    }
    catch {
        Write-ActivityLog "Error showing demo analytics: $($_.Exception.Message)" -Level "Error"
    }
}

function Open-SitesDeepDive {
    <#
    .SYNOPSIS
    Opens the Sites deep dive window
    #>
    try {
        Write-ActivityLog "Opening Sites deep dive from Visual Analytics" -Level "Information"
        
        # Check if we have data
        $sites = Get-SharePointData -DataType "Sites"
        
        if ($sites.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No sites data available. Please run 'Get All Sites' first.",
                "No Data",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            return
        }
        
        # Load the deep dive module if not already loaded
        $deepDivePath = Join-Path $PSScriptRoot "DeepDive\SitesDeepDive.ps1"
        if (Test-Path $deepDivePath) {
            . $deepDivePath
        }
        else {
            # Try alternative path
            $deepDivePath = Join-Path (Split-Path -Parent $PSScriptRoot) "UI\DeepDive\SitesDeepDive.ps1"
            if (Test-Path $deepDivePath) {
                . $deepDivePath
            }
        }
        
        # Show the deep dive window
        Show-SitesDeepDive
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Open-SitesDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Sites deep dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}
function Open-UsersDeepDive {
    <#
    .SYNOPSIS
    Opens the Users deep dive window
    #>
    try {
        Write-ActivityLog "Opening Users deep dive from Visual Analytics" -Level "Information"

        $users = Get-SharePointData -DataType "Users"

        if ($users.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No users data available. Please run 'Get Users' first.",
                "No Data",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            return
        }

        Show-UsersDeepDive
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Open-UsersDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Users deep dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Open-GroupsDeepDive {
    <#
    .SYNOPSIS
    Opens the Groups deep dive window
    #>
    try {
        Write-ActivityLog "Opening Groups deep dive from Visual Analytics" -Level "Information"

        $groups = Get-SharePointData -DataType "Groups"

        if ($groups.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No groups data available. Please run 'Get Groups' first.",
                "No Data",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            return
        }

        Show-GroupsDeepDive
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Open-GroupsDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Groups deep dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}
function Open-ExternalUsersDeepDive {
    <#
    .SYNOPSIS
    Opens the External Users deep dive window
    #>
    try {
        Write-ActivityLog "Opening External Users deep dive from Visual Analytics" -Level "Information"

        $users = Get-SharePointData -DataType "Users"

        if ($users.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No users data available. Please run 'Get Users' first.",
                "No Data",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            return
        }

        Show-ExternalUsersDeepDive
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Open-ExternalUsersDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open External Users deep dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}
function Open-PermissionsDeepDive {
    <#
    .SYNOPSIS
    Opens the Role Assignment Mapping deep dive window
    #>
    try {
        Write-ActivityLog "Opening Permissions deep dive from Visual Analytics" -Level "Information"

        $roleAssignments = Get-SharePointData -DataType "RoleAssignments"

        if ($roleAssignments.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No role assignment data available. Please run 'Analyze Permissions' first.",
                "No Data",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            return
        }

        Show-PermissionsDeepDive
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Open-PermissionsDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Permissions deep dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Open-InheritanceDeepDive {
    <#
    .SYNOPSIS
    Opens the Permission Inheritance Analysis deep dive window
    #>
    try {
        Write-ActivityLog "Opening Inheritance deep dive from Visual Analytics" -Level "Information"

        $items = Get-SharePointData -DataType "InheritanceItems"

        if ($items.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No inheritance data available. Please run 'Analyze Permissions' first.",
                "No Data",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            return
        }

        Show-InheritanceDeepDive
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Open-InheritanceDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Inheritance deep dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Open-SharingLinksDeepDive {
    <#
    .SYNOPSIS
    Opens the Sharing Links Security Audit deep dive window
    #>
    try {
        Write-ActivityLog "Opening Sharing Links deep dive from Visual Analytics" -Level "Information"

        $links = Get-SharePointData -DataType "SharingLinks"

        if ($links.Count -eq 0) {
            [System.Windows.MessageBox]::Show(
                "No sharing links data available. Please run 'Analyze Permissions' first.",
                "No Data",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
            return
        }

        Show-SharingLinksDeepDive
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Open-SharingLinksDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Sharing Links deep dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Add-ClickableTooltips {
    <#
    .SYNOPSIS
    Adds tooltips to metric tiles to indicate they're clickable
    #>
    try {
        if ($script:txtTotalSites) {
            [System.Windows.Controls.ToolTipService]::SetToolTip(
                $script:txtTotalSites.Parent.Parent,
                "Click for detailed sites analysis"
            )
        }
        
        if ($script:txtTotalUsers) {
            [System.Windows.Controls.ToolTipService]::SetToolTip(
                $script:txtTotalUsers.Parent.Parent,
                "Click for detailed users analysis"
            )
        }
        
        if ($script:txtTotalGroups) {
            [System.Windows.Controls.ToolTipService]::SetToolTip(
                $script:txtTotalGroups.Parent.Parent,
                "Click for detailed groups analysis"
            )
        }
        
        if ($script:txtExternalUsers) {
            [System.Windows.Controls.ToolTipService]::SetToolTip(
                $script:txtExternalUsers.Parent.Parent,
                "Click for external users analysis"
            )
        }
        
    }
    catch {
        Write-ActivityLog "Error adding tooltips: $($_.Exception.Message)" -Level "Warning"
    }
}