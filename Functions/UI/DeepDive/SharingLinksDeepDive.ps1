# ============================================
# SharingLinksDeepDive.ps1 - Sharing Links Security Audit
# ============================================
# Location: Functions/UI/DeepDive/SharingLinksDeepDive.ps1

function Show-SharingLinksDeepDive {
    <#
    .SYNOPSIS
    Shows the Sharing Links Security Audit deep dive window
    #>
    try {
        Write-ActivityLog "Opening Sharing Links Deep Dive window" -Level "Information"

        # Load XAML
        $xamlPath = Join-Path $PSScriptRoot "..\..\..\Views\DeepDive\SharingLinksDeepDive.xaml"
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
            txtLinkCount = $deepDiveWindow.FindName("txtLinkCount")
            btnRefreshData = $deepDiveWindow.FindName("btnRefreshData")
            btnExport = $deepDiveWindow.FindName("btnExport")
            # Summary Stats
            txtTotalLinks = $deepDiveWindow.FindName("txtTotalLinks")
            txtAnonymousLinks = $deepDiveWindow.FindName("txtAnonymousLinks")
            txtOrganizationLinks = $deepDiveWindow.FindName("txtOrganizationLinks")
            txtSpecificLinks = $deepDiveWindow.FindName("txtSpecificLinks")
            txtTotalRecipients = $deepDiveWindow.FindName("txtTotalRecipients")
            # Sharing Links Tab
            txtSearch = $deepDiveWindow.FindName("txtSearch")
            cboLinkTypeFilter = $deepDiveWindow.FindName("cboLinkTypeFilter")
            cboAccessFilter = $deepDiveWindow.FindName("cboAccessFilter")
            dgSharingLinks = $deepDiveWindow.FindName("dgSharingLinks")
            # Link Distribution Tab
            canvasLinkDistribution = $deepDiveWindow.FindName("canvasLinkDistribution")
            dgLinkBreakdown = $deepDiveWindow.FindName("dgLinkBreakdown")
            # Security Findings Tab
            txtHighRiskLinks = $deepDiveWindow.FindName("txtHighRiskLinks")
            txtMediumRiskLinks = $deepDiveWindow.FindName("txtMediumRiskLinks")
            txtLowRiskLinks = $deepDiveWindow.FindName("txtLowRiskLinks")
            lstSecurityFindings = $deepDiveWindow.FindName("lstSecurityFindings")
            # Status Bar
            txtStatus = $deepDiveWindow.FindName("txtStatus")
            txtLastUpdate = $deepDiveWindow.FindName("txtLastUpdate")
        }

        # Set up event handlers
        $controls.btnRefreshData.Add_Click({
            Refresh-SharingLinksDeepDiveData -Controls $controls
        })

        $controls.btnExport.Add_Click({
            Export-SharingLinksDeepDiveData -Controls $controls
        })

        $controls.txtSearch.Add_TextChanged({
            Apply-SharingLinksFilter -Controls $controls
        })

        $controls.cboLinkTypeFilter.Add_SelectionChanged({
            Apply-SharingLinksFilter -Controls $controls
        })

        $controls.cboAccessFilter.Add_SelectionChanged({
            Apply-SharingLinksFilter -Controls $controls
        })

        # Load initial data
        Load-SharingLinksDeepDiveData -Controls $controls

        # Show window
        $deepDiveWindow.ShowDialog() | Out-Null

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Show-SharingLinksDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Sharing Links Deep Dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Load-SharingLinksDeepDiveData {
    <#
    .SYNOPSIS
    Loads data into the Sharing Links Deep Dive window
    #>
    param($Controls)

    try {
        $Controls.txtStatus.Text = "Loading sharing links data..."

        $links = Get-SharePointData -DataType "SharingLinks"

        # Update header
        $Controls.txtLinkCount.Text = "Auditing $($links.Count) sharing links"

        # Calculate summary statistics
        $anonymousCount = 0
        $organizationCount = 0
        $specificCount = 0
        $totalRecipients = 0

        foreach ($link in $links) {
            $linkType = if ($link["LinkType"]) { $link["LinkType"] } else { "Unknown" }
            switch ($linkType) {
                "Anonymous" { $anonymousCount++ }
                "Company-wide" { $organizationCount++ }
                "Organization" { $organizationCount++ }
                "Specific People" { $specificCount++ }
                default { $specificCount++ }
            }
            $memberCount = if ($link["MemberCount"]) { [int]$link["MemberCount"] } else { 0 }
            $totalRecipients += $memberCount
        }

        $Controls.txtTotalLinks.Text = $links.Count.ToString()
        $Controls.txtAnonymousLinks.Text = $anonymousCount.ToString()
        $Controls.txtOrganizationLinks.Text = $organizationCount.ToString()
        $Controls.txtSpecificLinks.Text = $specificCount.ToString()
        $Controls.txtTotalRecipients.Text = $totalRecipients.ToString()

        # Load grid
        Load-SharingLinksGrid -Controls $Controls -Links $links

        # Load distribution chart
        Load-SharingLinkDistribution -Controls $Controls -Links $links

        # Load security findings
        Load-SharingLinksSecurityFindings -Controls $Controls -Links $links

        $Controls.txtStatus.Text = "Ready"
        $Controls.txtLastUpdate.Text = "Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-SharingLinksDeepDiveData"
        $Controls.txtStatus.Text = "Error loading data"
    }
}

function Load-SharingLinksGrid {
    <#
    .SYNOPSIS
    Loads the sharing links data grid
    #>
    param($Controls, $Links)

    try {
        $linkObjects = @()

        foreach ($link in $Links) {
            $linkObjects += [PSCustomObject]@{
                GroupName = if ($link["GroupName"]) { $link["GroupName"] } else { "Unknown" }
                LinkType = if ($link["LinkType"]) { $link["LinkType"] } else { "Unknown" }
                AccessLevel = if ($link["AccessLevel"]) { $link["AccessLevel"] } else { "Unknown" }
                MemberCount = if ($link["MemberCount"]) { $link["MemberCount"] } else { 0 }
                SiteTitle = if ($link["SiteTitle"]) { $link["SiteTitle"] } else { "N/A" }
                CreatedDate = if ($link["CreatedDate"]) { $link["CreatedDate"] } else { "N/A" }
            }
        }

        $Controls.dgSharingLinks.ItemsSource = $linkObjects

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-SharingLinksGrid"
    }
}

function Load-SharingLinkDistribution {
    <#
    .SYNOPSIS
    Loads sharing link distribution chart and breakdown
    #>
    param($Controls, $Links)

    try {
        $Controls.canvasLinkDistribution.Children.Clear()

        # Count by link type
        $typeCounts = @{}
        $typeRecipients = @{}
        foreach ($link in $Links) {
            $linkType = if ($link["LinkType"]) { $link["LinkType"] } else { "Unknown" }
            $members = if ($link["MemberCount"]) { [int]$link["MemberCount"] } else { 0 }

            if ($typeCounts.ContainsKey($linkType)) {
                $typeCounts[$linkType]++
                $typeRecipients[$linkType] += $members
            } else {
                $typeCounts[$linkType] = 1
                $typeRecipients[$linkType] = $members
            }
        }

        if ($typeCounts.Count -eq 0) { return }

        $total = ($typeCounts.Values | Measure-Object -Sum).Sum
        if ($total -eq 0) { $total = 1 }

        $colorMap = @{
            "Anonymous"       = "#DC3545"
            "Company-wide"    = "#FFC107"
            "Organization"    = "#FFC107"
            "Specific People" = "#28A745"
            "Unknown"         = "#6C757D"
        }

        # Build breakdown data
        $breakdownData = @()
        foreach ($key in $typeCounts.Keys | Sort-Object { $typeCounts[$_] } -Descending) {
            $count = $typeCounts[$key]
            $percentage = [math]::Round(($count / $total) * 100, 1)
            $color = if ($colorMap.ContainsKey($key)) { $colorMap[$key] } else { "#6F42C1" }
            $recipients = if ($typeRecipients.ContainsKey($key)) { $typeRecipients[$key] } else { 0 }

            $breakdownData += [PSCustomObject]@{
                LinkType = $key
                Count = $count
                Percentage = "$percentage%"
                TotalRecipients = $recipients
                PercentageValue = $percentage
                BarColor = $color
            }
        }

        $Controls.dgLinkBreakdown.ItemsSource = $breakdownData

        # Draw chart
        $startY = 20
        $barHeight = 30
        $maxBarWidth = 180

        for ($i = 0; $i -lt $breakdownData.Count; $i++) {
            $item = $breakdownData[$i]
            $barWidth = [Math]::Max(($item.PercentageValue / 100) * $maxBarWidth, 5)

            $colorHex = $item.BarColor.TrimStart('#').ToUpper()
            if ($colorHex.Length -eq 6 -and $colorHex -match '^[0-9A-F]{6}$') {
                $r = [System.Convert]::ToByte($colorHex.Substring(0, 2), 16)
                $g = [System.Convert]::ToByte($colorHex.Substring(2, 2), 16)
                $b = [System.Convert]::ToByte($colorHex.Substring(4, 2), 16)
                $brush = [System.Windows.Media.SolidColorBrush]::new(
                    [System.Windows.Media.Color]::FromArgb(255, $r, $g, $b))
            } else {
                $brush = [System.Windows.Media.SolidColorBrush]::new(
                    [System.Windows.Media.Colors]::Gray)
            }

            $bar = New-Object System.Windows.Shapes.Rectangle
            $bar.Width = $barWidth
            $bar.Height = $barHeight
            $bar.Fill = $brush

            [System.Windows.Controls.Canvas]::SetLeft($bar, 10)
            [System.Windows.Controls.Canvas]::SetTop($bar, $startY + ($i * ($barHeight + 15)))
            $Controls.canvasLinkDistribution.Children.Add($bar)

            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = "$($item.LinkType): $($item.Count) links, $($item.TotalRecipients) recipients"
            $label.FontSize = 10
            $label.Foreground = "#495057"

            [System.Windows.Controls.Canvas]::SetLeft($label, $barWidth + 20)
            [System.Windows.Controls.Canvas]::SetTop($label, $startY + ($i * ($barHeight + 15)) + 8)
            $Controls.canvasLinkDistribution.Children.Add($label)
        }

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-SharingLinkDistribution"
    }
}

function Load-SharingLinksSecurityFindings {
    <#
    .SYNOPSIS
    Loads security findings for sharing links
    #>
    param($Controls, $Links)

    try {
        $highRisk = 0
        $mediumRisk = 0
        $lowRisk = 0
        $findings = @()

        # Check for anonymous links (highest risk)
        $anonymousLinks = $Links | Where-Object { $_["LinkType"] -eq "Anonymous" }
        if ($anonymousLinks.Count -gt 0) {
            $highRisk++
            $anonymousEdit = ($anonymousLinks | Where-Object { $_["AccessLevel"] -in @("Edit", "Full Control") }).Count
            $findings += [PSCustomObject]@{
                Icon = "🚨"
                Finding = "Anonymous Sharing Links Detected"
                Detail = "$($anonymousLinks.Count) anonymous sharing links found ($anonymousEdit with edit access)"
                Recommendation = "Review all anonymous links. Remove edit access from anonymous links. Consider disabling anonymous sharing at the tenant level."
                SeverityColor = "#FFCDD2"
            }
        }

        # Check for anonymous links with edit access (critical)
        $anonymousEdit = $Links | Where-Object {
            $_["LinkType"] -eq "Anonymous" -and $_["AccessLevel"] -in @("Edit", "Full Control")
        }
        if ($anonymousEdit.Count -gt 0) {
            $highRisk++
            $findings += [PSCustomObject]@{
                Icon = "🔴"
                Finding = "Anonymous Edit Access"
                Detail = "$($anonymousEdit.Count) anonymous links grant edit or full control access"
                Recommendation = "CRITICAL: Anonymous edit links allow anyone with the link to modify content. Remove immediately unless explicitly required."
                SeverityColor = "#FFCDD2"
            }
        }

        # Check for company-wide links (medium risk)
        $orgLinks = $Links | Where-Object { $_["LinkType"] -in @("Company-wide", "Organization") }
        if ($orgLinks.Count -gt 5) {
            $mediumRisk++
            $findings += [PSCustomObject]@{
                Icon = "⚠️"
                Finding = "Excessive Company-wide Sharing"
                Detail = "$($orgLinks.Count) company-wide sharing links detected"
                Recommendation = "Company-wide links expose content to all employees. Use specific-people links for sensitive content."
                SeverityColor = "#FFE0B2"
            }
        } elseif ($orgLinks.Count -gt 0) {
            $lowRisk++
            $findings += [PSCustomObject]@{
                Icon = "ℹ️"
                Finding = "Company-wide Sharing Links"
                Detail = "$($orgLinks.Count) company-wide sharing links in use"
                Recommendation = "Verify company-wide links are appropriate for the shared content."
                SeverityColor = "#E1F5FE"
            }
        }

        # Check for high recipient counts
        $highRecipientLinks = $Links | Where-Object {
            $_["MemberCount"] -and [int]$_["MemberCount"] -gt 20
        }
        if ($highRecipientLinks.Count -gt 0) {
            $mediumRisk++
            $findings += [PSCustomObject]@{
                Icon = "⚠️"
                Finding = "High Recipient Count Links"
                Detail = "$($highRecipientLinks.Count) links shared with more than 20 recipients"
                Recommendation = "Links with many recipients may indicate oversharing. Consider using groups or site-level permissions instead."
                SeverityColor = "#FFE0B2"
            }
        }

        # Positive findings
        if ($anonymousLinks.Count -eq 0) {
            $findings += [PSCustomObject]@{
                Icon = "✅"
                Finding = "No Anonymous Links"
                Detail = "No anonymous sharing links detected - good security posture"
                Recommendation = "Continue monitoring for newly created anonymous links."
                SeverityColor = "#C8E6C9"
            }
        }

        $specificLinks = $Links | Where-Object { $_["LinkType"] -eq "Specific People" }
        if ($specificLinks.Count -gt 0 -and $anonymousLinks.Count -eq 0 -and $orgLinks.Count -eq 0) {
            $findings += [PSCustomObject]@{
                Icon = "✅"
                Finding = "Targeted Sharing Only"
                Detail = "All $($specificLinks.Count) sharing links target specific people"
                Recommendation = "Excellent security practice. Continue using specific-people links."
                SeverityColor = "#C8E6C9"
            }
        }

        $findings += [PSCustomObject]@{
            Icon = "📊"
            Finding = "Audit Summary"
            Detail = "Analyzed $($Links.Count) sharing links across all sites"
            Recommendation = "Schedule regular sharing link audits to prevent unauthorized access spread."
            SeverityColor = "#E1F5FE"
        }

        $Controls.txtHighRiskLinks.Text = $highRisk.ToString()
        $Controls.txtMediumRiskLinks.Text = $mediumRisk.ToString()
        $Controls.txtLowRiskLinks.Text = $lowRisk.ToString()
        $Controls.lstSecurityFindings.ItemsSource = $findings

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-SharingLinksSecurityFindings"
    }
}

function Apply-SharingLinksFilter {
    <#
    .SYNOPSIS
    Applies filters to the sharing links data grid
    #>
    param($Controls)

    try {
        $searchText = $Controls.txtSearch.Text.ToLower()
        $linkTypeFilter = $Controls.cboLinkTypeFilter.SelectedItem.Content
        $accessFilter = $Controls.cboAccessFilter.SelectedItem.Content

        $links = Get-SharePointData -DataType "SharingLinks"
        $filtered = @()

        foreach ($link in $links) {
            $include = $true

            # Search filter
            if ($searchText -and $searchText.Length -gt 0) {
                $name = if ($link["GroupName"]) { $link["GroupName"].ToLower() } else { "" }
                $site = if ($link["SiteTitle"]) { $link["SiteTitle"].ToLower() } else { "" }
                if (-not ($name.Contains($searchText) -or $site.Contains($searchText))) {
                    $include = $false
                }
            }

            # Link type filter
            if ($linkTypeFilter -ne "All Link Types" -and $include) {
                $lt = if ($link["LinkType"]) { $link["LinkType"] } else { "" }
                if ($lt -ne $linkTypeFilter) { $include = $false }
            }

            # Access level filter
            if ($accessFilter -ne "All Access Levels" -and $include) {
                $al = if ($link["AccessLevel"]) { $link["AccessLevel"] } else { "" }
                if ($al -ne $accessFilter) { $include = $false }
            }

            if ($include) { $filtered += $link }
        }

        Load-SharingLinksGrid -Controls $Controls -Links $filtered

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Apply-SharingLinksFilter"
    }
}

function Refresh-SharingLinksDeepDiveData {
    <#
    .SYNOPSIS
    Refreshes the sharing links deep dive data
    #>
    param($Controls)

    try {
        $Controls.txtStatus.Text = "Refreshing data..."

        Load-SharingLinksDeepDiveData -Controls $Controls

        [System.Windows.MessageBox]::Show(
            "Data refreshed successfully!",
            "Success",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Refresh-SharingLinksDeepDiveData"
        $Controls.txtStatus.Text = "Error refreshing data"
    }
}

function Export-SharingLinksDeepDiveData {
    <#
    .SYNOPSIS
    Exports the sharing links data to CSV
    #>
    param($Controls)

    try {
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "SharePoint_SharingLinks_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        if ($saveDialog.ShowDialog() -eq $true) {
            $links = Get-SharePointData -DataType "SharingLinks"

            $exportData = @()
            foreach ($link in $links) {
                $exportData += [PSCustomObject]@{
                    GroupName = $link["GroupName"]
                    LinkType = $link["LinkType"]
                    AccessLevel = $link["AccessLevel"]
                    MemberCount = $link["MemberCount"]
                    SiteTitle = $link["SiteTitle"]
                    CreatedDate = $link["CreatedDate"]
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
        Write-ErrorLog -Message $_.Exception.Message -Location "Export-SharingLinksDeepDiveData"
        [System.Windows.MessageBox]::Show(
            "Failed to export data: $($_.Exception.Message)",
            "Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}
