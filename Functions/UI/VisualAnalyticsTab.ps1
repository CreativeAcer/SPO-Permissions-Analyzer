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
        
        # Initialize with empty state
        Reset-VisualAnalytics
        
        Write-ActivityLog "Visual Analytics tab initialized successfully" -Level "Information"
    }
    catch {
        Write-ActivityLog "Failed to initialize Visual Analytics tab: $($_.Exception.Message)" -Level "Error"
        throw
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
    Updates the permission distribution chart
    #>
    try {
        $script:canvasPermissionChart.Children.Clear()
        
        # Demo permission distribution data - can be made dynamic later
        $permissions = @(
            @{Name = "Full Control"; Count = 12; Color = "#DC3545"},
            @{Name = "Edit"; Count = 25; Color = "#FFC107"},
            @{Name = "Read"; Count = 18; Color = "#28A745"},
            @{Name = "View Only"; Count = 5; Color = "#17A2B8"}
        )
        
        $total = ($permissions | Measure-Object -Property Count -Sum).Sum
        if ($total -eq 0) { $total = 1 } # Prevent division by zero
        
        $startY = 20
        $barHeight = 25
        $maxBarWidth = 150
        
        for ($i = 0; $i -lt $permissions.Count; $i++) {
            $perm = $permissions[$i]
            $barWidth = ($perm.Count / $total) * $maxBarWidth
            
            # Create permission bar
            $bar = New-Object System.Windows.Shapes.Rectangle
            $bar.Width = $barWidth
            $bar.Height = $barHeight
            $bar.Fill = $perm.Color
            
            # Position bar
            [System.Windows.Controls.Canvas]::SetLeft($bar, 10)
            [System.Windows.Controls.Canvas]::SetTop($bar, $startY + ($i * ($barHeight + 10)))
            
            $script:canvasPermissionChart.Children.Add($bar)
            
            # Add permission label
            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = "$($perm.Name): $($perm.Count)"
            $label.FontSize = 10
            $label.Foreground = "#495057"
            
            [System.Windows.Controls.Canvas]::SetLeft($label, $barWidth + 20)
            [System.Windows.Controls.Canvas]::SetTop($label, $startY + ($i * ($barHeight + 10)) + 5)
            
            $script:canvasPermissionChart.Children.Add($label)
        }
        
        Write-ActivityLog "Permission chart updated successfully" -Level "Information"
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