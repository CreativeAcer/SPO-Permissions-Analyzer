# ============================================
# InheritanceDeepDive.ps1 - Permission Inheritance Analysis
# ============================================
# Location: Functions/UI/DeepDive/InheritanceDeepDive.ps1

function Show-InheritanceDeepDive {
    <#
    .SYNOPSIS
    Shows the Permission Inheritance Analysis deep dive window
    #>
    try {
        Write-ActivityLog "Opening Inheritance Deep Dive window" -Level "Information"

        # Load XAML
        $xamlPath = Join-Path $PSScriptRoot "..\..\..\Views\DeepDive\InheritanceDeepDive.xaml"
        if (-not (Test-Path $xamlPath)) {
            throw "Deep Dive XAML file not found at: $xamlPath"
        }

        $xamlContent = Get-Content $xamlPath -Raw
        $reader = [System.Xml.XmlNodeReader]::new([xml]$xamlContent)
        $deepDiveWindow = [System.Windows.Markup.XamlReader]::Load($reader)

        # Get controls
        $controls = @{
            Window = $deepDiveWindow
            # Header
            txtItemCount = $deepDiveWindow.FindName("txtItemCount")
            btnRefreshData = $deepDiveWindow.FindName("btnRefreshData")
            btnExport = $deepDiveWindow.FindName("btnExport")
            # Summary Stats
            txtTotalItems = $deepDiveWindow.FindName("txtTotalItems")
            txtInheritingCount = $deepDiveWindow.FindName("txtInheritingCount")
            txtBrokenCount = $deepDiveWindow.FindName("txtBrokenCount")
            txtListCount = $deepDiveWindow.FindName("txtListCount")
            txtLibraryCount = $deepDiveWindow.FindName("txtLibraryCount")
            # Inheritance Tree Tab
            txtSearch = $deepDiveWindow.FindName("txtSearch")
            cboTypeFilter = $deepDiveWindow.FindName("cboTypeFilter")
            cboInheritanceFilter = $deepDiveWindow.FindName("cboInheritanceFilter")
            dgInheritanceItems = $deepDiveWindow.FindName("dgInheritanceItems")
            # Inheritance Overview Tab
            canvasInheritanceChart = $deepDiveWindow.FindName("canvasInheritanceChart")
            dgInheritanceSummary = $deepDiveWindow.FindName("dgInheritanceSummary")
            # Findings Tab
            txtHealthyItems = $deepDiveWindow.FindName("txtHealthyItems")
            txtWarningItems = $deepDiveWindow.FindName("txtWarningItems")
            txtCriticalItems = $deepDiveWindow.FindName("txtCriticalItems")
            lstInheritanceFindings = $deepDiveWindow.FindName("lstInheritanceFindings")
            # Status Bar
            txtStatus = $deepDiveWindow.FindName("txtStatus")
            txtLastUpdate = $deepDiveWindow.FindName("txtLastUpdate")
        }

        # Set up event handlers
        $controls.btnRefreshData.Add_Click({
            Refresh-InheritanceDeepDiveData -Controls $controls
        })

        $controls.btnExport.Add_Click({
            Export-InheritanceDeepDiveData -Controls $controls
        })

        $controls.txtSearch.Add_TextChanged({
            Apply-InheritanceFilter -Controls $controls
        })

        $controls.cboTypeFilter.Add_SelectionChanged({
            Apply-InheritanceFilter -Controls $controls
        })

        $controls.cboInheritanceFilter.Add_SelectionChanged({
            Apply-InheritanceFilter -Controls $controls
        })

        # Load initial data
        Load-InheritanceDeepDiveData -Controls $controls

        # Show window
        $deepDiveWindow.ShowDialog() | Out-Null

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Show-InheritanceDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Inheritance Deep Dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Load-InheritanceDeepDiveData {
    <#
    .SYNOPSIS
    Loads data into the Inheritance Deep Dive window
    #>
    param($Controls)

    try {
        $Controls.txtStatus.Text = "Loading inheritance data..."

        $items = Get-SharePointData -DataType "InheritanceItems"

        # Update header
        $Controls.txtItemCount.Text = "Analyzing $($items.Count) items"

        # Calculate summary statistics
        $inheritingCount = 0
        $brokenCount = 0
        $listCount = 0
        $libraryCount = 0

        foreach ($item in $items) {
            if ($item["HasUniquePermissions"] -eq $true) {
                $brokenCount++
            } else {
                $inheritingCount++
            }

            $type = if ($item["Type"]) { $item["Type"] } else { "Unknown" }
            switch ($type) {
                "List" { $listCount++ }
                "Document Library" { $libraryCount++ }
                "Library" { $libraryCount++ }
            }
        }

        $Controls.txtTotalItems.Text = $items.Count.ToString()
        $Controls.txtInheritingCount.Text = $inheritingCount.ToString()
        $Controls.txtBrokenCount.Text = $brokenCount.ToString()
        $Controls.txtListCount.Text = $listCount.ToString()
        $Controls.txtLibraryCount.Text = $libraryCount.ToString()

        # Load grid
        Load-InheritanceGrid -Controls $Controls -Items $items

        # Load overview chart
        Load-InheritanceOverview -Controls $Controls -Items $items

        # Load findings
        Load-InheritanceFindings -Controls $Controls -Items $items

        $Controls.txtStatus.Text = "Ready"
        $Controls.txtLastUpdate.Text = "Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-InheritanceDeepDiveData"
        $Controls.txtStatus.Text = "Error loading data"
    }
}

function Load-InheritanceGrid {
    <#
    .SYNOPSIS
    Loads the inheritance items data grid
    #>
    param($Controls, $Items)

    try {
        $gridObjects = @()

        foreach ($item in $Items) {
            $gridObjects += [PSCustomObject]@{
                Title = if ($item["Title"]) { $item["Title"] } else { "Unknown" }
                Type = if ($item["Type"]) { $item["Type"] } else { "Unknown" }
                Url = if ($item["Url"]) { $item["Url"] } else { "N/A" }
                HasUniquePermissions = if ($item["HasUniquePermissions"]) { $item["HasUniquePermissions"] } else { $false }
                ParentUrl = if ($item["ParentUrl"]) { $item["ParentUrl"] } else { "N/A" }
                RoleAssignmentCount = if ($item["RoleAssignmentCount"]) { $item["RoleAssignmentCount"] } else { 0 }
                SiteTitle = if ($item["SiteTitle"]) { $item["SiteTitle"] } else { "N/A" }
            }
        }

        $Controls.dgInheritanceItems.ItemsSource = $gridObjects

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-InheritanceGrid"
    }
}

function Load-InheritanceOverview {
    <#
    .SYNOPSIS
    Loads inheritance overview chart and summary
    #>
    param($Controls, $Items)

    try {
        $Controls.canvasInheritanceChart.Children.Clear()

        if ($Items.Count -eq 0) { return }

        # Group by site
        $siteSummary = @{}
        foreach ($item in $Items) {
            $siteTitle = if ($item["SiteTitle"]) { $item["SiteTitle"] } else { "Unknown Site" }
            if (-not $siteSummary.ContainsKey($siteTitle)) {
                $siteSummary[$siteTitle] = @{ Total = 0; Inheriting = 0; Broken = 0 }
            }
            $siteSummary[$siteTitle].Total++
            if ($item["HasUniquePermissions"] -eq $true) {
                $siteSummary[$siteTitle].Broken++
            } else {
                $siteSummary[$siteTitle].Inheriting++
            }
        }

        # Build summary grid data
        $summaryData = @()
        foreach ($key in $siteSummary.Keys | Sort-Object) {
            $data = $siteSummary[$key]
            $breakPct = if ($data.Total -gt 0) { [math]::Round(($data.Broken / $data.Total) * 100, 1) } else { 0 }
            $summaryData += [PSCustomObject]@{
                SiteTitle = $key
                TotalItems = $data.Total
                InheritingItems = $data.Inheriting
                BrokenItems = $data.Broken
                BreakPercentage = "$breakPct%"
                PercentageValue = $breakPct
                BarColor = if ($breakPct -gt 50) { "#DC3545" } elseif ($breakPct -gt 25) { "#FFC107" } else { "#28A745" }
            }
        }

        $Controls.dgInheritanceSummary.ItemsSource = $summaryData

        # Draw simple chart: inheriting vs broken
        $totalInheriting = ($Items | Where-Object { $_["HasUniquePermissions"] -ne $true }).Count
        $totalBroken = ($Items | Where-Object { $_["HasUniquePermissions"] -eq $true }).Count
        $total = $Items.Count
        if ($total -eq 0) { $total = 1 }

        # Inheriting bar
        $inheritWidth = [Math]::Max(($totalInheriting / $total) * 250, 5)
        $inheritBar = New-Object System.Windows.Shapes.Rectangle
        $inheritBar.Width = $inheritWidth
        $inheritBar.Height = 40
        $inheritBar.Fill = "#28A745"
        [System.Windows.Controls.Canvas]::SetLeft($inheritBar, 10)
        [System.Windows.Controls.Canvas]::SetTop($inheritBar, 30)
        $Controls.canvasInheritanceChart.Children.Add($inheritBar)

        $inheritLabel = New-Object System.Windows.Controls.TextBlock
        $inheritLabel.Text = "Inheriting: $totalInheriting ($([math]::Round(($totalInheriting / $total) * 100, 1))%)"
        $inheritLabel.FontSize = 11
        $inheritLabel.Foreground = "#495057"
        [System.Windows.Controls.Canvas]::SetLeft($inheritLabel, $inheritWidth + 20)
        [System.Windows.Controls.Canvas]::SetTop($inheritLabel, 40)
        $Controls.canvasInheritanceChart.Children.Add($inheritLabel)

        # Broken bar
        $brokenWidth = [Math]::Max(($totalBroken / $total) * 250, 5)
        $brokenBar = New-Object System.Windows.Shapes.Rectangle
        $brokenBar.Width = $brokenWidth
        $brokenBar.Height = 40
        $brokenBar.Fill = "#DC3545"
        [System.Windows.Controls.Canvas]::SetLeft($brokenBar, 10)
        [System.Windows.Controls.Canvas]::SetTop($brokenBar, 90)
        $Controls.canvasInheritanceChart.Children.Add($brokenBar)

        $brokenLabel = New-Object System.Windows.Controls.TextBlock
        $brokenLabel.Text = "Broken Inheritance: $totalBroken ($([math]::Round(($totalBroken / $total) * 100, 1))%)"
        $brokenLabel.FontSize = 11
        $brokenLabel.Foreground = "#495057"
        [System.Windows.Controls.Canvas]::SetLeft($brokenLabel, $brokenWidth + 20)
        [System.Windows.Controls.Canvas]::SetTop($brokenLabel, 100)
        $Controls.canvasInheritanceChart.Children.Add($brokenLabel)

        # Title
        $titleLabel = New-Object System.Windows.Controls.TextBlock
        $titleLabel.Text = "Inheritance Status Distribution"
        $titleLabel.FontSize = 12
        $titleLabel.FontWeight = "Bold"
        $titleLabel.Foreground = "#495057"
        [System.Windows.Controls.Canvas]::SetLeft($titleLabel, 10)
        [System.Windows.Controls.Canvas]::SetTop($titleLabel, 5)
        $Controls.canvasInheritanceChart.Children.Add($titleLabel)

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-InheritanceOverview"
    }
}

function Load-InheritanceFindings {
    <#
    .SYNOPSIS
    Loads inheritance-related findings and recommendations
    #>
    param($Controls, $Items)

    try {
        $healthyCount = 0
        $warningCount = 0
        $criticalCount = 0
        $findings = @()

        $brokenItems = $Items | Where-Object { $_["HasUniquePermissions"] -eq $true }
        $totalItems = $Items.Count

        # Check overall break percentage
        if ($totalItems -gt 0) {
            $breakPercentage = [math]::Round(($brokenItems.Count / $totalItems) * 100, 1)

            if ($breakPercentage -gt 50) {
                $criticalCount++
                $findings += [PSCustomObject]@{
                    Icon = "🚨"
                    Site = "Tenant-wide"
                    Issue = "$breakPercentage% of items have broken inheritance ($($brokenItems.Count)/$totalItems)"
                    Recommendation = "Excessive permission breaks create management overhead. Consider consolidating permissions at site level."
                    SeverityColor = "#FFCDD2"
                }
            } elseif ($breakPercentage -gt 25) {
                $warningCount++
                $findings += [PSCustomObject]@{
                    Icon = "⚠️"
                    Site = "Tenant-wide"
                    Issue = "$breakPercentage% of items have broken inheritance ($($brokenItems.Count)/$totalItems)"
                    Recommendation = "Review broken inheritance items and consolidate where possible."
                    SeverityColor = "#FFE0B2"
                }
            } else {
                $healthyCount++
            }
        }

        # Check for document libraries with broken inheritance (high impact)
        $brokenLibraries = $brokenItems | Where-Object {
            $_["Type"] -eq "Document Library" -or $_["Type"] -eq "Library"
        }
        if ($brokenLibraries.Count -gt 0) {
            $warningCount++
            $findings += [PSCustomObject]@{
                Icon = "📁"
                Site = "Multiple Libraries"
                Issue = "$($brokenLibraries.Count) document libraries have unique permissions"
                Recommendation = "Document libraries with unique permissions may confuse users. Verify access is intentional."
                SeverityColor = "#FFE0B2"
            }
        }

        # Check for items with many role assignments (complexity indicator)
        $complexItems = $Items | Where-Object {
            $_["RoleAssignmentCount"] -and [int]$_["RoleAssignmentCount"] -gt 10
        }
        if ($complexItems.Count -gt 0) {
            $warningCount++
            $findings += [PSCustomObject]@{
                Icon = "⚠️"
                Site = "Multiple locations"
                Issue = "$($complexItems.Count) items have more than 10 role assignments"
                Recommendation = "High role assignment counts indicate overly complex permissions. Consider using groups instead."
                SeverityColor = "#FFE0B2"
            }
        }

        # Positive finding
        if ($findings.Count -eq 0) {
            $healthyCount++
            $findings += [PSCustomObject]@{
                Icon = "✅"
                Site = "Tenant-wide"
                Issue = "Permission inheritance structure looks healthy"
                Recommendation = "Continue monitoring for inheritance breaks during content restructuring."
                SeverityColor = "#C8E6C9"
            }
        }

        $findings += [PSCustomObject]@{
            Icon = "📊"
            Site = "Summary"
            Issue = "Analyzed $totalItems items: $($Items.Count - $brokenItems.Count) inheriting, $($brokenItems.Count) with unique permissions"
            Recommendation = "Regular inheritance audits help maintain a clean permission structure."
            SeverityColor = "#E1F5FE"
        }

        $Controls.txtHealthyItems.Text = $healthyCount.ToString()
        $Controls.txtWarningItems.Text = $warningCount.ToString()
        $Controls.txtCriticalItems.Text = $criticalCount.ToString()
        $Controls.lstInheritanceFindings.ItemsSource = $findings

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-InheritanceFindings"
    }
}

function Apply-InheritanceFilter {
    <#
    .SYNOPSIS
    Applies filters to the inheritance items data grid
    #>
    param($Controls)

    try {
        $searchText = $Controls.txtSearch.Text.ToLower()
        $typeFilter = $Controls.cboTypeFilter.SelectedItem.Content
        $inheritanceFilter = $Controls.cboInheritanceFilter.SelectedItem.Content

        $items = Get-SharePointData -DataType "InheritanceItems"
        $filtered = @()

        foreach ($item in $items) {
            $include = $true

            # Search filter
            if ($searchText -and $searchText.Length -gt 0) {
                $title = if ($item["Title"]) { $item["Title"].ToLower() } else { "" }
                $url = if ($item["Url"]) { $item["Url"].ToLower() } else { "" }
                if (-not ($title.Contains($searchText) -or $url.Contains($searchText))) {
                    $include = $false
                }
            }

            # Type filter
            if ($typeFilter -ne "All Types" -and $include) {
                $itemType = if ($item["Type"]) { $item["Type"] } else { "" }
                if ($itemType -ne $typeFilter) { $include = $false }
            }

            # Inheritance filter
            if ($inheritanceFilter -ne "All Items" -and $include) {
                $hasUnique = $item["HasUniquePermissions"] -eq $true
                switch ($inheritanceFilter) {
                    "Inheriting" { if ($hasUnique) { $include = $false } }
                    "Broken Inheritance" { if (-not $hasUnique) { $include = $false } }
                }
            }

            if ($include) { $filtered += $item }
        }

        Load-InheritanceGrid -Controls $Controls -Items $filtered

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Apply-InheritanceFilter"
    }
}

function Refresh-InheritanceDeepDiveData {
    <#
    .SYNOPSIS
    Refreshes the inheritance deep dive data
    #>
    param($Controls)

    try {
        $Controls.txtStatus.Text = "Refreshing data..."

        Load-InheritanceDeepDiveData -Controls $Controls

        [System.Windows.MessageBox]::Show(
            "Data refreshed successfully!",
            "Success",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Refresh-InheritanceDeepDiveData"
        $Controls.txtStatus.Text = "Error refreshing data"
    }
}

function Export-InheritanceDeepDiveData {
    <#
    .SYNOPSIS
    Exports the inheritance data to CSV
    #>
    param($Controls)

    try {
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "SharePoint_Inheritance_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        if ($saveDialog.ShowDialog() -eq $true) {
            $items = Get-SharePointData -DataType "InheritanceItems"

            $exportData = @()
            foreach ($item in $items) {
                $exportData += [PSCustomObject]@{
                    Title = $item["Title"]
                    Type = $item["Type"]
                    Url = $item["Url"]
                    HasUniquePermissions = $item["HasUniquePermissions"]
                    ParentUrl = $item["ParentUrl"]
                    RoleAssignmentCount = $item["RoleAssignmentCount"]
                    SiteTitle = $item["SiteTitle"]
                }
            }

            $exportData | Export-Csv -Path $saveDialog.FileName -NoTypeInformation

            [System.Windows.MessageBox]::Show(
                "Data exported successfully to:`n$($saveDialog.FileName)",
                "Export Complete",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )

            $Controls.txtStatus.Text = "Data exported successfully"
        }

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Export-InheritanceDeepDiveData"
        [System.Windows.MessageBox]::Show(
            "Failed to export data: $($_.Exception.Message)",
            "Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}
