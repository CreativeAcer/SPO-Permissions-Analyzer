# ============================================
# GroupsDeepDive.ps1 - Groups Deep Dive Window
# ============================================
# Location: Functions/UI/DeepDive/GroupsDeepDive.ps1

function Show-GroupsDeepDive {
    <#
    .SYNOPSIS
    Shows the Groups Deep Dive window with detailed analysis
    #>
    try {
        Write-ActivityLog "Opening Groups Deep Dive window" -Level "Information"

        # Load XAML
        $xamlPath = Join-Path $PSScriptRoot "..\..\..\Views\DeepDive\GroupsDeepDive.xaml"
        if (-not (Test-Path $xamlPath)) {
            throw "Groups Deep Dive XAML file not found at: $xamlPath"
        }

        $xamlContent = Get-Content $xamlPath -Raw
        $reader = [System.Xml.XmlNodeReader]::new([xml]$xamlContent)
        $deepDiveWindow = [System.Windows.Markup.XamlReader]::Load($reader)

        # Get controls
        $controls = @{
            Window = $deepDiveWindow
            # Header
            txtGroupCount = $deepDiveWindow.FindName("txtGroupCount")
            btnRefreshData = $deepDiveWindow.FindName("btnRefreshData")
            btnExport = $deepDiveWindow.FindName("btnExport")
            # Summary Stats
            txtTotalGroups = $deepDiveWindow.FindName("txtTotalGroups")
            txtTotalMembers = $deepDiveWindow.FindName("txtTotalMembers")
            txtAvgMembers = $deepDiveWindow.FindName("txtAvgMembers")
            txtEmptyGroups = $deepDiveWindow.FindName("txtEmptyGroups")
            txtLargeGroups = $deepDiveWindow.FindName("txtLargeGroups")
            # All Groups Tab
            txtSearch = $deepDiveWindow.FindName("txtSearch")
            cboSizeFilter = $deepDiveWindow.FindName("cboSizeFilter")
            dgGroups = $deepDiveWindow.FindName("dgGroups")
            # Membership Analysis Tab
            canvasMemberDistribution = $deepDiveWindow.FindName("canvasMemberDistribution")
            dgLargestGroups = $deepDiveWindow.FindName("dgLargestGroups")
            # Group Health Tab
            txtHealthyGroups = $deepDiveWindow.FindName("txtHealthyGroups")
            txtWarningGroups = $deepDiveWindow.FindName("txtWarningGroups")
            txtCriticalGroups = $deepDiveWindow.FindName("txtCriticalGroups")
            lstGroupIssues = $deepDiveWindow.FindName("lstGroupIssues")
            # Status Bar
            txtStatus = $deepDiveWindow.FindName("txtStatus")
            txtLastUpdate = $deepDiveWindow.FindName("txtLastUpdate")
        }

        # Set up event handlers
        $controls.btnRefreshData.Add_Click({
            Refresh-GroupsDeepDiveData -Controls $controls
        })

        $controls.btnExport.Add_Click({
            Export-GroupsDeepDiveData -Controls $controls
        })

        $controls.txtSearch.Add_TextChanged({
            Apply-GroupsFilter -Controls $controls
        })

        $controls.cboSizeFilter.Add_SelectionChanged({
            Apply-GroupsFilter -Controls $controls
        })

        # Load initial data
        Load-GroupsDeepDiveData -Controls $controls

        # Show window
        $deepDiveWindow.ShowDialog() | Out-Null

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Show-GroupsDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Groups Deep Dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Load-GroupsDeepDiveData {
    <#
    .SYNOPSIS
    Loads data into the Groups Deep Dive window
    #>
    param($Controls)

    try {
        $Controls.txtStatus.Text = "Loading groups data..."

        $groups = Get-SharePointData -DataType "Groups"

        if ($groups.Count -eq 0) {
            $Controls.txtStatus.Text = "No group data available. Run 'Get Groups' first."
            return
        }

        # Build group objects and calculate stats
        $groupObjects = @()
        $totalMembers = 0
        $emptyCount = 0
        $largeCount = 0

        foreach ($group in $groups) {
            $name = if ($group["Name"]) { $group["Name"] } else { "Unknown" }
            $owner = if ($group["Owner"]) { $group["Owner"] } else { "N/A" }
            $memberCount = if ($group["MemberCount"]) { [int]$group["MemberCount"] } else { 0 }
            $permission = if ($group["Permission"]) { $group["Permission"] } else { "Unknown" }
            $site = if ($group["Site"]) { $group["Site"] } else { "N/A" }
            $groupType = if ($group["GroupType"]) { $group["GroupType"] } else { "SharePoint" }

            $totalMembers += $memberCount
            if ($memberCount -eq 0) { $emptyCount++ }
            if ($memberCount -ge 10) { $largeCount++ }

            $groupObjects += [PSCustomObject]@{
                GroupName   = $name
                Owner       = $owner
                MemberCount = $memberCount
                Permission  = $permission
                Site        = $site
                GroupType   = $groupType
            }
        }

        $avgMembers = if ($groups.Count -gt 0) { [math]::Round($totalMembers / $groups.Count, 1) } else { 0 }

        # Update header and summary stats
        $Controls.txtGroupCount.Text = "Analyzing $($groups.Count) groups"
        $Controls.txtTotalGroups.Text = $groups.Count.ToString()
        $Controls.txtTotalMembers.Text = $totalMembers.ToString()
        $Controls.txtAvgMembers.Text = $avgMembers.ToString()
        $Controls.txtEmptyGroups.Text = $emptyCount.ToString()
        $Controls.txtLargeGroups.Text = $largeCount.ToString()

        # Populate grid
        $Controls.dgGroups.ItemsSource = $groupObjects

        # Load membership analysis
        Load-GroupMembershipAnalysis -Controls $Controls -Groups $groups -TotalMembers $totalMembers

        # Load group health
        Load-GroupHealthAnalysis -Controls $Controls -Groups $groups

        $Controls.txtStatus.Text = "Ready"
        $Controls.txtLastUpdate.Text = "Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-GroupsDeepDiveData"
        $Controls.txtStatus.Text = "Error loading data"
    }
}

function Load-GroupMembershipAnalysis {
    <#
    .SYNOPSIS
    Loads membership analysis chart and largest groups grid
    #>
    param($Controls, $Groups, $TotalMembers)

    try {
        $Controls.canvasMemberDistribution.Children.Clear()

        # Sort groups by member count descending
        $sorted = $Groups | Sort-Object { if ($_["MemberCount"]) { [int]$_["MemberCount"] } else { 0 } } -Descending

        # Build largest groups data
        $largestData = @()
        $rank = 1
        foreach ($group in ($sorted | Select-Object -First 10)) {
            $memberCount = if ($group["MemberCount"]) { [int]$group["MemberCount"] } else { 0 }
            $pct = if ($TotalMembers -gt 0) { [math]::Round(($memberCount / $TotalMembers) * 100, 1) } else { 0 }

            $color = switch ($true) {
                ($memberCount -ge 20) { "#DC3545" }
                ($memberCount -ge 10) { "#FFC107" }
                ($memberCount -ge 5) { "#17A2B8" }
                default { "#28A745" }
            }

            $largestData += [PSCustomObject]@{
                Rank            = $rank
                GroupName       = if ($group["Name"]) { $group["Name"] } else { "Unknown" }
                MemberCount     = $memberCount
                Percentage      = "$pct%"
                PercentageValue = $pct
                BarColor        = $color
            }
            $rank++
        }

        $Controls.dgLargestGroups.ItemsSource = $largestData

        # Draw size distribution chart (buckets: 0, 1-5, 6-10, 10+)
        $buckets = @(
            @{ Label = "Empty (0)"; Count = 0; Color = "#6C757D" },
            @{ Label = "Small (1-5)"; Count = 0; Color = "#28A745" },
            @{ Label = "Medium (6-10)"; Count = 0; Color = "#17A2B8" },
            @{ Label = "Large (10+)"; Count = 0; Color = "#DC3545" }
        )

        foreach ($group in $Groups) {
            $mc = if ($group["MemberCount"]) { [int]$group["MemberCount"] } else { 0 }
            switch ($true) {
                ($mc -eq 0)  { $buckets[0].Count++ }
                ($mc -le 5)  { $buckets[1].Count++ }
                ($mc -le 10) { $buckets[2].Count++ }
                default      { $buckets[3].Count++ }
            }
        }

        $maxBucket = ($buckets | Measure-Object -Property Count -Maximum).Maximum
        if ($maxBucket -eq 0) { $maxBucket = 1 }

        $startY = 20
        $barHeight = 30
        $maxBarWidth = 200

        for ($i = 0; $i -lt $buckets.Count; $i++) {
            $bucket = $buckets[$i]
            $barWidth = [Math]::Max(($bucket.Count / $maxBucket) * $maxBarWidth, 5)

            $bar = New-Object System.Windows.Shapes.Rectangle
            $bar.Width = $barWidth
            $bar.Height = $barHeight
            $bar.Fill = $bucket.Color

            [System.Windows.Controls.Canvas]::SetLeft($bar, 10)
            [System.Windows.Controls.Canvas]::SetTop($bar, $startY + ($i * ($barHeight + 10)))
            $Controls.canvasMemberDistribution.Children.Add($bar)

            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = "$($bucket.Label): $($bucket.Count)"
            $label.FontSize = 10
            $label.Foreground = "#495057"

            [System.Windows.Controls.Canvas]::SetLeft($label, $barWidth + 20)
            [System.Windows.Controls.Canvas]::SetTop($label, $startY + ($i * ($barHeight + 10)) + 7)
            $Controls.canvasMemberDistribution.Children.Add($label)
        }
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-GroupMembershipAnalysis"
    }
}

function Load-GroupHealthAnalysis {
    <#
    .SYNOPSIS
    Loads group health analysis
    #>
    param($Controls, $Groups)

    try {
        $healthy = 0
        $warning = 0
        $critical = 0
        $issues = @()

        foreach ($group in $Groups) {
            $name = if ($group["Name"]) { $group["Name"] } else { "Unknown" }
            $memberCount = if ($group["MemberCount"]) { [int]$group["MemberCount"] } else { 0 }
            $owner = if ($group["Owner"]) { $group["Owner"] } else { "" }
            $hasIssue = $false

            # Empty group check
            if ($memberCount -eq 0) {
                $warning++
                $hasIssue = $true
                $issues += [PSCustomObject]@{
                    Group = $name
                    Issue = "Empty group with no members"
                    Recommendation = "Remove empty groups or add appropriate members"
                    Icon = "⚠️"
                    SeverityColor = "#FFE0B2"
                }
            }

            # Very large group check
            if ($memberCount -gt 25) {
                $critical++
                $hasIssue = $true
                $issues += [PSCustomObject]@{
                    Group = $name
                    Issue = "Very large group with $memberCount members"
                    Recommendation = "Consider breaking into smaller, role-based groups for better governance"
                    Icon = "🚨"
                    SeverityColor = "#FFCDD2"
                }
            }
            elseif ($memberCount -gt 10 -and -not $hasIssue) {
                $warning++
                $hasIssue = $true
                $issues += [PSCustomObject]@{
                    Group = $name
                    Issue = "Large group with $memberCount members"
                    Recommendation = "Review membership regularly to ensure all members still need access"
                    Icon = "📋"
                    SeverityColor = "#E1F5FE"
                }
            }

            # No owner check
            if (-not $owner -or $owner -eq "N/A" -or $owner -eq "") {
                if (-not $hasIssue) { $warning++ }
                $hasIssue = $true
                $issues += [PSCustomObject]@{
                    Group = $name
                    Issue = "Group has no designated owner"
                    Recommendation = "Assign an owner for accountability and access reviews"
                    Icon = "👤"
                    SeverityColor = "#FFE0B2"
                }
            }

            if (-not $hasIssue) { $healthy++ }
        }

        $Controls.txtHealthyGroups.Text = $healthy.ToString()
        $Controls.txtWarningGroups.Text = $warning.ToString()
        $Controls.txtCriticalGroups.Text = $critical.ToString()
        $Controls.lstGroupIssues.ItemsSource = $issues
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-GroupHealthAnalysis"
    }
}

function Apply-GroupsFilter {
    <#
    .SYNOPSIS
    Applies filters to the groups data grid
    #>
    param($Controls)

    try {
        $searchText = $Controls.txtSearch.Text.ToLower()
        $sizeFilter = $Controls.cboSizeFilter.SelectedItem.Content

        $groups = Get-SharePointData -DataType "Groups"
        $filtered = @()

        foreach ($group in $groups) {
            $include = $true
            $name = if ($group["Name"]) { $group["Name"] } else { "" }
            $owner = if ($group["Owner"]) { $group["Owner"] } else { "" }
            $memberCount = if ($group["MemberCount"]) { [int]$group["MemberCount"] } else { 0 }
            $permission = if ($group["Permission"]) { $group["Permission"] } else { "Unknown" }
            $site = if ($group["Site"]) { $group["Site"] } else { "" }
            $groupType = if ($group["GroupType"]) { $group["GroupType"] } else { "SharePoint" }

            # Search filter
            if ($searchText -and $searchText.Length -gt 0) {
                if (-not ($name.ToLower().Contains($searchText) -or $owner.ToLower().Contains($searchText) -or $site.ToLower().Contains($searchText))) {
                    $include = $false
                }
            }

            # Size filter
            if ($include -and $sizeFilter -ne "All Sizes") {
                switch ($sizeFilter) {
                    "Empty (0 members)"     { if ($memberCount -ne 0) { $include = $false } }
                    "Small (1-5 members)"   { if ($memberCount -lt 1 -or $memberCount -gt 5) { $include = $false } }
                    "Medium (6-10 members)" { if ($memberCount -lt 6 -or $memberCount -gt 10) { $include = $false } }
                    "Large (10+ members)"   { if ($memberCount -lt 10) { $include = $false } }
                }
            }

            if ($include) {
                $filtered += [PSCustomObject]@{
                    GroupName   = $name
                    Owner       = $owner
                    MemberCount = $memberCount
                    Permission  = $permission
                    Site        = $site
                    GroupType   = $groupType
                }
            }
        }

        $Controls.dgGroups.ItemsSource = $filtered
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Apply-GroupsFilter"
    }
}

function Refresh-GroupsDeepDiveData {
    <#
    .SYNOPSIS
    Refreshes group data
    #>
    param($Controls)

    try {
        if (-not $script:SPOConnected) {
            [System.Windows.MessageBox]::Show(
                "Not connected to SharePoint. Please connect first.",
                "Warning",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Warning
            )
            return
        }

        $Controls.txtStatus.Text = "Refreshing group data..."
        Load-GroupsDeepDiveData -Controls $Controls

        [System.Windows.MessageBox]::Show(
            "Group data refreshed successfully!",
            "Success",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Refresh-GroupsDeepDiveData"
        $Controls.txtStatus.Text = "Error refreshing data"
    }
}

function Export-GroupsDeepDiveData {
    <#
    .SYNOPSIS
    Exports group deep dive data to CSV
    #>
    param($Controls)

    try {
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "SharePoint_Groups_DeepDive_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        if ($saveDialog.ShowDialog() -eq $true) {
            $groups = Get-SharePointData -DataType "Groups"

            $exportData = @()
            foreach ($group in $groups) {
                $exportData += [PSCustomObject]@{
                    GroupName   = $group["Name"]
                    Owner       = $group["Owner"]
                    MemberCount = $group["MemberCount"]
                    Permission  = $group["Permission"]
                    Site        = $group["Site"]
                    GroupType   = $group["GroupType"]
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
        Write-ErrorLog -Message $_.Exception.Message -Location "Export-GroupsDeepDiveData"
        [System.Windows.MessageBox]::Show(
            "Failed to export data: $($_.Exception.Message)",
            "Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}
