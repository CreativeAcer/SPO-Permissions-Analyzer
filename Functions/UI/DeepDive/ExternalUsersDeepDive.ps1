# ============================================
# ExternalUsersDeepDive.ps1 - External Users Deep Dive Window
# ============================================
# Location: Functions/UI/DeepDive/ExternalUsersDeepDive.ps1

function Show-ExternalUsersDeepDive {
    <#
    .SYNOPSIS
    Shows the External Users Deep Dive window with security-focused analysis
    #>
    try {
        Write-ActivityLog "Opening External Users Deep Dive window" -Level "Information"

        # Load XAML
        $xamlPath = Join-Path $PSScriptRoot "..\..\..\Views\DeepDive\ExternalUsersDeepDive.xaml"
        if (-not (Test-Path $xamlPath)) {
            throw "External Users Deep Dive XAML file not found at: $xamlPath"
        }

        $xamlContent = Get-Content $xamlPath -Raw
        $reader = [System.Xml.XmlNodeReader]::new([xml]$xamlContent)
        $deepDiveWindow = [System.Windows.Markup.XamlReader]::Load($reader)

        # Get controls
        $controls = @{
            Window = $deepDiveWindow
            # Header
            txtExternalCount = $deepDiveWindow.FindName("txtExternalCount")
            btnRefreshData = $deepDiveWindow.FindName("btnRefreshData")
            btnExport = $deepDiveWindow.FindName("btnExport")
            # Summary Stats
            txtTotalExternal = $deepDiveWindow.FindName("txtTotalExternal")
            txtExternalDomains = $deepDiveWindow.FindName("txtExternalDomains")
            txtSitesAccessed = $deepDiveWindow.FindName("txtSitesAccessed")
            txtEditAccess = $deepDiveWindow.FindName("txtEditAccess")
            # External Users Tab
            txtSearch = $deepDiveWindow.FindName("txtSearch")
            cboDomainFilter = $deepDiveWindow.FindName("cboDomainFilter")
            dgExternalUsers = $deepDiveWindow.FindName("dgExternalUsers")
            # Domain Analysis Tab
            canvasDomainChart = $deepDiveWindow.FindName("canvasDomainChart")
            dgDomains = $deepDiveWindow.FindName("dgDomains")
            # Security Audit Tab
            txtReadOnly = $deepDiveWindow.FindName("txtReadOnly")
            txtCanEdit = $deepDiveWindow.FindName("txtCanEdit")
            txtFullControl = $deepDiveWindow.FindName("txtFullControl")
            lstAuditFindings = $deepDiveWindow.FindName("lstAuditFindings")
            # Status Bar
            txtStatus = $deepDiveWindow.FindName("txtStatus")
            txtLastUpdate = $deepDiveWindow.FindName("txtLastUpdate")
        }

        # Set up event handlers
        $controls.btnRefreshData.Add_Click({
            Refresh-ExternalUsersDeepDiveData -Controls $controls
        })

        $controls.btnExport.Add_Click({
            Export-ExternalUsersDeepDiveData -Controls $controls
        })

        $controls.txtSearch.Add_TextChanged({
            Apply-ExternalUsersFilter -Controls $controls
        })

        $controls.cboDomainFilter.Add_SelectionChanged({
            Apply-ExternalUsersFilter -Controls $controls
        })

        # Load initial data
        Load-ExternalUsersDeepDiveData -Controls $controls

        # Show window
        $deepDiveWindow.ShowDialog() | Out-Null

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Show-ExternalUsersDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open External Users Deep Dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Get-ExternalUsersFromData {
    <#
    .SYNOPSIS
    Extracts external users from the users data store
    #>
    $users = Get-SharePointData -DataType "Users"
    $external = @()

    foreach ($user in $users) {
        $isExternal = $false
        $email = if ($user["Email"]) { $user["Email"] } else { "" }

        if ($user["IsExternal"]) {
            $isExternal = [bool]$user["IsExternal"]
        } elseif ($email -match "#ext#|guest|_.*@.*\.onmicrosoft\.com") {
            $isExternal = $true
        }

        if ($isExternal) {
            $external += $user
        }
    }

    return $external
}

function Load-ExternalUsersDeepDiveData {
    <#
    .SYNOPSIS
    Loads data into the External Users Deep Dive window
    #>
    param($Controls)

    try {
        $Controls.txtStatus.Text = "Loading external users data..."

        $externalUsers = Get-ExternalUsersFromData

        if ($externalUsers.Count -eq 0) {
            $Controls.txtExternalCount.Text = "No external users found"
            $Controls.txtTotalExternal.Text = "0"
            $Controls.txtExternalDomains.Text = "0"
            $Controls.txtSitesAccessed.Text = "0"
            $Controls.txtEditAccess.Text = "0"
            $Controls.txtStatus.Text = "No external users detected - good security posture"

            # Show positive audit finding
            $Controls.lstAuditFindings.ItemsSource = @(
                [PSCustomObject]@{
                    Finding = "No External Users Detected"
                    Detail = "No external or guest users have access to tenant resources"
                    Recommendation = "Continue monitoring for new external sharing"
                    Icon = "✅"
                    SeverityColor = "#28A745"
                }
            )
            return
        }

        # Build user objects and compute stats
        $userObjects = @()
        $domainCounts = @{}
        $sitesSet = @{}
        $editCount = 0

        foreach ($user in $externalUsers) {
            $name = if ($user["Name"]) { $user["Name"] } else { "Unknown" }
            $email = if ($user["Email"]) { $user["Email"] } else { "N/A" }
            $permission = if ($user["Permission"]) { $user["Permission"] } else { "Unknown" }
            $site = if ($user["Site"]) { $user["Site"] } else { "N/A" }

            # Extract domain from email
            $domain = "Unknown"
            if ($email -match "@(.+)$") {
                $domain = $Matches[1]
            }

            if (-not $domainCounts.ContainsKey($domain)) { $domainCounts[$domain] = 0 }
            $domainCounts[$domain]++

            if ($site -ne "N/A") { $sitesSet[$site] = $true }

            if ($permission -match "Edit|Full Control|Contribute") { $editCount++ }

            $userObjects += [PSCustomObject]@{
                Name       = $name
                Email      = $email
                Domain     = $domain
                Permission = $permission
                Site       = $site
            }
        }

        # Update summary stats
        $Controls.txtExternalCount.Text = "Analyzing $($externalUsers.Count) external users"
        $Controls.txtTotalExternal.Text = $externalUsers.Count.ToString()
        $Controls.txtExternalDomains.Text = $domainCounts.Count.ToString()
        $Controls.txtSitesAccessed.Text = $sitesSet.Count.ToString()
        $Controls.txtEditAccess.Text = $editCount.ToString()

        # Populate grid
        $Controls.dgExternalUsers.ItemsSource = $userObjects

        # Populate domain filter dropdown
        $Controls.cboDomainFilter.Items.Clear()
        $allItem = New-Object System.Windows.Controls.ComboBoxItem
        $allItem.Content = "All Domains"
        $allItem.IsSelected = $true
        $Controls.cboDomainFilter.Items.Add($allItem) | Out-Null
        foreach ($domain in $domainCounts.Keys | Sort-Object) {
            $item = New-Object System.Windows.Controls.ComboBoxItem
            $item.Content = $domain
            $Controls.cboDomainFilter.Items.Add($item) | Out-Null
        }

        # Load domain analysis
        Load-DomainAnalysis -Controls $Controls -DomainCounts $domainCounts -TotalExternal $externalUsers.Count

        # Load security audit
        Load-ExternalSecurityAudit -Controls $Controls -ExternalUsers $externalUsers -DomainCounts $domainCounts -EditCount $editCount

        $Controls.txtStatus.Text = "Ready"
        $Controls.txtLastUpdate.Text = "Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-ExternalUsersDeepDiveData"
        $Controls.txtStatus.Text = "Error loading data"
    }
}

function Load-DomainAnalysis {
    <#
    .SYNOPSIS
    Loads domain analysis chart and grid
    #>
    param($Controls, $DomainCounts, $TotalExternal)

    try {
        $Controls.canvasDomainChart.Children.Clear()

        # Build domain data sorted by count
        $domainData = @()
        $domainColors = @("#DC3545", "#FFC107", "#17A2B8", "#28A745", "#6F42C1", "#FD7E14", "#20C997", "#6C757D")
        $colorIndex = 0

        foreach ($domain in $DomainCounts.Keys | Sort-Object { $DomainCounts[$_] } -Descending) {
            $count = $DomainCounts[$domain]
            $pct = if ($TotalExternal -gt 0) { [math]::Round(($count / $TotalExternal) * 100, 1) } else { 0 }
            $color = $domainColors[$colorIndex % $domainColors.Count]
            $colorIndex++

            $domainData += [PSCustomObject]@{
                Domain          = $domain
                UserCount       = $count
                Percentage      = "$pct%"
                PercentageValue = $pct
                BarColor        = $color
            }
        }

        $Controls.dgDomains.ItemsSource = $domainData

        # Draw chart
        $maxCount = if ($domainData.Count -gt 0) { ($domainData | Measure-Object -Property UserCount -Maximum).Maximum } else { 1 }
        if ($maxCount -eq 0) { $maxCount = 1 }

        $startY = 20
        $barHeight = 25
        $maxBarWidth = 200

        for ($i = 0; $i -lt [Math]::Min($domainData.Count, 6); $i++) {
            $item = $domainData[$i]
            $barWidth = [Math]::Max(($item.UserCount / $maxCount) * $maxBarWidth, 5)

            $bar = New-Object System.Windows.Shapes.Rectangle
            $bar.Width = $barWidth
            $bar.Height = $barHeight
            $bar.Fill = $item.BarColor

            [System.Windows.Controls.Canvas]::SetLeft($bar, 10)
            [System.Windows.Controls.Canvas]::SetTop($bar, $startY + ($i * ($barHeight + 10)))
            $Controls.canvasDomainChart.Children.Add($bar)

            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = "$($item.Domain): $($item.UserCount)"
            $label.FontSize = 10
            $label.Foreground = "#495057"

            [System.Windows.Controls.Canvas]::SetLeft($label, $barWidth + 20)
            [System.Windows.Controls.Canvas]::SetTop($label, $startY + ($i * ($barHeight + 10)) + 5)
            $Controls.canvasDomainChart.Children.Add($label)
        }
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-DomainAnalysis"
    }
}

function Load-ExternalSecurityAudit {
    <#
    .SYNOPSIS
    Loads external user security audit findings
    #>
    param($Controls, $ExternalUsers, $DomainCounts, $EditCount)

    try {
        $readOnly = 0
        $canEdit = 0
        $fullControl = 0

        foreach ($user in $ExternalUsers) {
            $permission = if ($user["Permission"]) { $user["Permission"] } else { "Unknown" }
            switch -Regex ($permission) {
                "Full Control" { $fullControl++ }
                "Edit|Contribute" { $canEdit++ }
                default { $readOnly++ }
            }
        }

        $Controls.txtReadOnly.Text = $readOnly.ToString()
        $Controls.txtCanEdit.Text = $canEdit.ToString()
        $Controls.txtFullControl.Text = $fullControl.ToString()

        # Generate audit findings
        $findings = @()

        if ($fullControl -gt 0) {
            $findings += [PSCustomObject]@{
                Finding = "External Users with Full Control"
                Detail = "$fullControl external users have Full Control access"
                Recommendation = "Immediately review and revoke Full Control from external users unless explicitly required"
                Icon = "🚨"
                SeverityColor = "#DC3545"
            }
        }

        if ($canEdit -gt 0) {
            $findings += [PSCustomObject]@{
                Finding = "External Users with Edit Access"
                Detail = "$canEdit external users can edit content"
                Recommendation = "Verify each external editor's need for write access; consider downgrading to Read"
                Icon = "⚠️"
                SeverityColor = "#FFC107"
            }
        }

        if ($DomainCounts.Count -gt 3) {
            $findings += [PSCustomObject]@{
                Finding = "Multiple External Domains"
                Detail = "$($DomainCounts.Count) different external domains have access"
                Recommendation = "Review external domains and ensure all are expected business partners"
                Icon = "🏢"
                SeverityColor = "#FFC107"
            }
        }

        foreach ($domain in $DomainCounts.Keys) {
            if ($DomainCounts[$domain] -gt 5) {
                $findings += [PSCustomObject]@{
                    Finding = "High User Count from $domain"
                    Detail = "$($DomainCounts[$domain]) users from $domain have access"
                    Recommendation = "Consider using a security group for $domain instead of individual user access"
                    Icon = "📊"
                    SeverityColor = "#17A2B8"
                }
            }
        }

        if ($readOnly -gt 0 -and $canEdit -eq 0 -and $fullControl -eq 0) {
            $findings += [PSCustomObject]@{
                Finding = "Read-Only External Access"
                Detail = "All $readOnly external users have read-only access"
                Recommendation = "Good security posture - external users cannot modify content"
                Icon = "✅"
                SeverityColor = "#28A745"
            }
        }

        $findings += [PSCustomObject]@{
            Finding = "Audit Complete"
            Detail = "Analyzed $($ExternalUsers.Count) external users across $($DomainCounts.Count) domains"
            Recommendation = "Schedule regular reviews of external user access"
            Icon = "📋"
            SeverityColor = "#6C757D"
        }

        $Controls.lstAuditFindings.ItemsSource = $findings
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-ExternalSecurityAudit"
    }
}

function Apply-ExternalUsersFilter {
    <#
    .SYNOPSIS
    Applies filters to the external users data grid
    #>
    param($Controls)

    try {
        $searchText = $Controls.txtSearch.Text.ToLower()
        $domainFilter = if ($Controls.cboDomainFilter.SelectedItem) { $Controls.cboDomainFilter.SelectedItem.Content } else { "All Domains" }

        $externalUsers = Get-ExternalUsersFromData
        $filtered = @()

        foreach ($user in $externalUsers) {
            $include = $true
            $name = if ($user["Name"]) { $user["Name"] } else { "" }
            $email = if ($user["Email"]) { $user["Email"] } else { "" }
            $permission = if ($user["Permission"]) { $user["Permission"] } else { "Unknown" }
            $site = if ($user["Site"]) { $user["Site"] } else { "" }

            $domain = "Unknown"
            if ($email -match "@(.+)$") { $domain = $Matches[1] }

            # Search filter
            if ($searchText -and $searchText.Length -gt 0) {
                if (-not ($name.ToLower().Contains($searchText) -or $email.ToLower().Contains($searchText) -or $domain.ToLower().Contains($searchText))) {
                    $include = $false
                }
            }

            # Domain filter
            if ($include -and $domainFilter -ne "All Domains") {
                if ($domain -ne $domainFilter) { $include = $false }
            }

            if ($include) {
                $filtered += [PSCustomObject]@{
                    Name       = $name
                    Email      = $email
                    Domain     = $domain
                    Permission = $permission
                    Site       = $site
                }
            }
        }

        $Controls.dgExternalUsers.ItemsSource = $filtered
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Apply-ExternalUsersFilter"
    }
}

function Refresh-ExternalUsersDeepDiveData {
    <#
    .SYNOPSIS
    Refreshes external user data
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

        $Controls.txtStatus.Text = "Refreshing external user data..."
        Load-ExternalUsersDeepDiveData -Controls $Controls

        [System.Windows.MessageBox]::Show(
            "External user data refreshed successfully!",
            "Success",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Refresh-ExternalUsersDeepDiveData"
        $Controls.txtStatus.Text = "Error refreshing data"
    }
}

function Export-ExternalUsersDeepDiveData {
    <#
    .SYNOPSIS
    Exports external users data to CSV
    #>
    param($Controls)

    try {
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "SharePoint_ExternalUsers_DeepDive_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        if ($saveDialog.ShowDialog() -eq $true) {
            $externalUsers = Get-ExternalUsersFromData

            $exportData = @()
            foreach ($user in $externalUsers) {
                $email = if ($user["Email"]) { $user["Email"] } else { "" }
                $domain = "Unknown"
                if ($email -match "@(.+)$") { $domain = $Matches[1] }

                $exportData += [PSCustomObject]@{
                    Name       = $user["Name"]
                    Email      = $email
                    Domain     = $domain
                    Permission = $user["Permission"]
                    Site       = $user["Site"]
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
        Write-ErrorLog -Message $_.Exception.Message -Location "Export-ExternalUsersDeepDiveData"
        [System.Windows.MessageBox]::Show(
            "Failed to export data: $($_.Exception.Message)",
            "Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}
