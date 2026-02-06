# ============================================
# UsersDeepDive.ps1 - Users Deep Dive Window
# ============================================
# Location: Functions/UI/DeepDive/UsersDeepDive.ps1

function Show-UsersDeepDive {
    <#
    .SYNOPSIS
    Shows the Users Deep Dive window with detailed analysis
    #>
    try {
        Write-ActivityLog "Opening Users Deep Dive window" -Level "Information"

        # Load XAML
        $xamlPath = Join-Path $PSScriptRoot "..\..\..\Views\DeepDive\UsersDeepDive.xaml"
        if (-not (Test-Path $xamlPath)) {
            throw "Users Deep Dive XAML file not found at: $xamlPath"
        }

        $xamlContent = Get-Content $xamlPath -Raw
        $reader = [System.Xml.XmlNodeReader]::new([xml]$xamlContent)
        $deepDiveWindow = [System.Windows.Markup.XamlReader]::Load($reader)

        # Get controls
        $controls = @{
            Window = $deepDiveWindow
            # Header
            txtUserCount = $deepDiveWindow.FindName("txtUserCount")
            btnRefreshData = $deepDiveWindow.FindName("btnRefreshData")
            btnExport = $deepDiveWindow.FindName("btnExport")
            # Summary Stats
            txtTotalUsers = $deepDiveWindow.FindName("txtTotalUsers")
            txtInternalUsers = $deepDiveWindow.FindName("txtInternalUsers")
            txtExternalUsers = $deepDiveWindow.FindName("txtExternalUsers")
            txtSiteAdmins = $deepDiveWindow.FindName("txtSiteAdmins")
            txtPermissionLevels = $deepDiveWindow.FindName("txtPermissionLevels")
            # All Users Tab
            txtSearch = $deepDiveWindow.FindName("txtSearch")
            cboTypeFilter = $deepDiveWindow.FindName("cboTypeFilter")
            cboPermissionFilter = $deepDiveWindow.FindName("cboPermissionFilter")
            dgUsers = $deepDiveWindow.FindName("dgUsers")
            # Permission Breakdown Tab
            canvasPermissionBreakdown = $deepDiveWindow.FindName("canvasPermissionBreakdown")
            dgPermissionBreakdown = $deepDiveWindow.FindName("dgPermissionBreakdown")
            # Security Alerts Tab
            txtLowRiskUsers = $deepDiveWindow.FindName("txtLowRiskUsers")
            txtMediumRiskUsers = $deepDiveWindow.FindName("txtMediumRiskUsers")
            txtHighRiskUsers = $deepDiveWindow.FindName("txtHighRiskUsers")
            lstSecurityAlerts = $deepDiveWindow.FindName("lstSecurityAlerts")
            # Status Bar
            txtStatus = $deepDiveWindow.FindName("txtStatus")
            txtLastUpdate = $deepDiveWindow.FindName("txtLastUpdate")
        }

        # Set up event handlers
        $controls.btnRefreshData.Add_Click({
            Refresh-UsersDeepDiveData -Controls $controls
        })

        $controls.btnExport.Add_Click({
            Export-UsersDeepDiveData -Controls $controls
        })

        $controls.txtSearch.Add_TextChanged({
            Apply-UsersFilter -Controls $controls
        })

        $controls.cboTypeFilter.Add_SelectionChanged({
            Apply-UsersFilter -Controls $controls
        })

        $controls.cboPermissionFilter.Add_SelectionChanged({
            Apply-UsersFilter -Controls $controls
        })

        # Load initial data
        Load-UsersDeepDiveData -Controls $controls

        # Show window
        $deepDiveWindow.ShowDialog() | Out-Null

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Show-UsersDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Users Deep Dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Load-UsersDeepDiveData {
    <#
    .SYNOPSIS
    Loads data into the Users Deep Dive window
    #>
    param($Controls)

    try {
        $Controls.txtStatus.Text = "Loading users data..."

        $users = Get-SharePointData -DataType "Users"

        if ($users.Count -eq 0) {
            $Controls.txtStatus.Text = "No user data available. Run 'Get Users' first."
            return
        }

        # Build user objects for the data grid
        $userObjects = @()
        $internalCount = 0
        $externalCount = 0
        $adminCount = 0
        $permissionSet = @{}

        foreach ($user in $users) {
            $name = if ($user["Name"]) { $user["Name"] } else { "Unknown" }
            $email = if ($user["Email"]) { $user["Email"] } else { "N/A" }
            $permission = if ($user["Permission"]) { $user["Permission"] } else { "Unknown" }
            $site = if ($user["Site"]) { $user["Site"] } else { "N/A" }
            $isSiteAdmin = if ($user["IsSiteAdmin"]) { [bool]$user["IsSiteAdmin"] } else { $false }

            # Determine user type
            $isExternal = $false
            if ($user["IsExternal"]) {
                $isExternal = [bool]$user["IsExternal"]
            } elseif ($email -match "#ext#|guest|_.*@.*\.onmicrosoft\.com") {
                $isExternal = $true
            }

            $userType = if ($isExternal) { "External" } else { "Internal" }

            if ($isExternal) { $externalCount++ } else { $internalCount++ }
            if ($isSiteAdmin) { $adminCount++ }
            if (-not $permissionSet.ContainsKey($permission)) { $permissionSet[$permission] = 0 }
            $permissionSet[$permission]++

            $userObjects += [PSCustomObject]@{
                Name        = $name
                Email       = $email
                UserType    = $userType
                Permission  = $permission
                Site        = $site
                IsSiteAdmin = $isSiteAdmin
            }
        }

        # Update header and summary stats
        $Controls.txtUserCount.Text = "Analyzing $($users.Count) users"
        $Controls.txtTotalUsers.Text = $users.Count.ToString()
        $Controls.txtInternalUsers.Text = $internalCount.ToString()
        $Controls.txtExternalUsers.Text = $externalCount.ToString()
        $Controls.txtSiteAdmins.Text = $adminCount.ToString()
        $Controls.txtPermissionLevels.Text = $permissionSet.Count.ToString()

        # Populate grid
        $Controls.dgUsers.ItemsSource = $userObjects

        # Load permission breakdown
        Load-UsersPermissionBreakdown -Controls $Controls -PermissionData $permissionSet -TotalUsers $users.Count

        # Load security analysis
        Load-UsersSecurityAnalysis -Controls $Controls -Users $users -ExternalCount $externalCount -AdminCount $adminCount

        $Controls.txtStatus.Text = "Ready"
        $Controls.txtLastUpdate.Text = "Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-UsersDeepDiveData"
        $Controls.txtStatus.Text = "Error loading data"
    }
}

function Load-UsersPermissionBreakdown {
    <#
    .SYNOPSIS
    Loads permission breakdown chart and grid
    #>
    param($Controls, $PermissionData, $TotalUsers)

    try {
        $Controls.canvasPermissionBreakdown.Children.Clear()

        if ($PermissionData.Count -eq 0) { return }

        $colorMap = @{
            "Full Control" = "#DC3545"; "Edit" = "#FFC107"; "Read" = "#28A745"
            "View Only" = "#17A2B8"; "Member" = "#FFC107"; "Unknown" = "#6C757D"
        }

        # Build breakdown data
        $breakdownData = @()
        $rank = 0
        foreach ($key in $PermissionData.Keys | Sort-Object { $PermissionData[$_] } -Descending) {
            $rank++
            $count = $PermissionData[$key]
            $pct = if ($TotalUsers -gt 0) { [math]::Round(($count / $TotalUsers) * 100, 1) } else { 0 }
            $color = if ($colorMap.ContainsKey($key)) { $colorMap[$key] } else { "#6C757D" }

            $breakdownData += [PSCustomObject]@{
                PermissionLevel = $key
                UserCount       = $count
                Percentage      = "$pct%"
                PercentageValue = $pct
                BarColor        = $color
            }
        }

        $Controls.dgPermissionBreakdown.ItemsSource = $breakdownData

        # Draw chart
        $startY = 20
        $barHeight = 25
        $maxBarWidth = 200
        $maxCount = ($PermissionData.Values | Measure-Object -Maximum).Maximum
        if ($maxCount -eq 0) { $maxCount = 1 }

        $i = 0
        foreach ($item in $breakdownData) {
            $barWidth = [Math]::Max(($item.PercentageValue / 100) * $maxBarWidth, 5)

            $bar = New-Object System.Windows.Shapes.Rectangle
            $bar.Width = $barWidth
            $bar.Height = $barHeight
            $bar.Fill = $item.BarColor

            [System.Windows.Controls.Canvas]::SetLeft($bar, 10)
            [System.Windows.Controls.Canvas]::SetTop($bar, $startY + ($i * ($barHeight + 10)))
            $Controls.canvasPermissionBreakdown.Children.Add($bar)

            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = "$($item.PermissionLevel): $($item.UserCount)"
            $label.FontSize = 10
            $label.Foreground = "#495057"

            [System.Windows.Controls.Canvas]::SetLeft($label, $barWidth + 20)
            [System.Windows.Controls.Canvas]::SetTop($label, $startY + ($i * ($barHeight + 10)) + 5)
            $Controls.canvasPermissionBreakdown.Children.Add($label)

            $i++
            if ($i -ge 5) { break }
        }
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-UsersPermissionBreakdown"
    }
}

function Load-UsersSecurityAnalysis {
    <#
    .SYNOPSIS
    Loads security analysis for users
    #>
    param($Controls, $Users, $ExternalCount, $AdminCount)

    try {
        $lowRisk = 0
        $mediumRisk = 0
        $highRisk = 0
        $findings = @()

        foreach ($user in $Users) {
            $permission = if ($user["Permission"]) { $user["Permission"] } else { "Unknown" }
            $isExternal = $false
            if ($user["IsExternal"]) { $isExternal = [bool]$user["IsExternal"] }
            elseif ($user["Email"] -match "#ext#|guest") { $isExternal = $true }
            $isSiteAdmin = if ($user["IsSiteAdmin"]) { [bool]$user["IsSiteAdmin"] } else { $false }

            # Risk classification
            if ($isExternal -and ($permission -match "Full Control|Edit")) {
                $highRisk++
            } elseif ($isSiteAdmin -or $isExternal -or $permission -match "Full Control") {
                $mediumRisk++
            } else {
                $lowRisk++
            }
        }

        $Controls.txtLowRiskUsers.Text = $lowRisk.ToString()
        $Controls.txtMediumRiskUsers.Text = $mediumRisk.ToString()
        $Controls.txtHighRiskUsers.Text = $highRisk.ToString()

        # Generate findings
        if ($ExternalCount -gt 0) {
            $findings += [PSCustomObject]@{
                Finding = "External Users Detected"
                Detail = "$ExternalCount external users have access to tenant resources"
                Recommendation = "Review each external user's access and remove any that are no longer needed"
                Icon = "🌐"
                SeverityColor = if ($ExternalCount -gt 5) { "#DC3545" } else { "#FFC107" }
            }
        }

        if ($AdminCount -gt 3) {
            $findings += [PSCustomObject]@{
                Finding = "High Number of Site Admins"
                Detail = "$AdminCount users have site administrator privileges"
                Recommendation = "Follow least-privilege principle - limit admin access to essential personnel only"
                Icon = "🔑"
                SeverityColor = "#FFC107"
            }
        }

        if ($highRisk -gt 0) {
            $findings += [PSCustomObject]@{
                Finding = "External Users with Edit/Full Control"
                Detail = "$highRisk external users have elevated permissions"
                Recommendation = "Immediately review external users with edit or full control access"
                Icon = "🚨"
                SeverityColor = "#DC3545"
            }
        }

        if ($ExternalCount -eq 0) {
            $findings += [PSCustomObject]@{
                Finding = "No External Users"
                Detail = "No external users detected in the tenant"
                Recommendation = "Good security posture - continue monitoring for new external access"
                Icon = "✅"
                SeverityColor = "#28A745"
            }
        }

        $findings += [PSCustomObject]@{
            Finding = "Analysis Complete"
            Detail = "Analyzed $($Users.Count) users: $lowRisk low risk, $mediumRisk medium risk, $highRisk high risk"
            Recommendation = "Review medium and high risk users regularly"
            Icon = "📊"
            SeverityColor = "#6C757D"
        }

        $Controls.lstSecurityAlerts.ItemsSource = $findings
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-UsersSecurityAnalysis"
    }
}

function Apply-UsersFilter {
    <#
    .SYNOPSIS
    Applies filters to the users data grid
    #>
    param($Controls)

    try {
        $searchText = $Controls.txtSearch.Text.ToLower()
        $typeFilter = $Controls.cboTypeFilter.SelectedItem.Content
        $permFilter = $Controls.cboPermissionFilter.SelectedItem.Content

        $users = Get-SharePointData -DataType "Users"
        $filtered = @()

        foreach ($user in $users) {
            $include = $true
            $name = if ($user["Name"]) { $user["Name"] } else { "" }
            $email = if ($user["Email"]) { $user["Email"] } else { "" }
            $permission = if ($user["Permission"]) { $user["Permission"] } else { "Unknown" }
            $site = if ($user["Site"]) { $user["Site"] } else { "" }
            $isSiteAdmin = if ($user["IsSiteAdmin"]) { [bool]$user["IsSiteAdmin"] } else { $false }

            $isExternal = $false
            if ($user["IsExternal"]) { $isExternal = [bool]$user["IsExternal"] }
            elseif ($email -match "#ext#|guest") { $isExternal = $true }
            $userType = if ($isExternal) { "External" } else { "Internal" }

            # Search filter
            if ($searchText -and $searchText.Length -gt 0) {
                if (-not ($name.ToLower().Contains($searchText) -or $email.ToLower().Contains($searchText) -or $site.ToLower().Contains($searchText))) {
                    $include = $false
                }
            }

            # Type filter
            if ($include -and $typeFilter -ne "All Users") {
                if ($typeFilter -eq "Internal Only" -and $isExternal) { $include = $false }
                if ($typeFilter -eq "External Only" -and -not $isExternal) { $include = $false }
            }

            # Permission filter
            if ($include -and $permFilter -ne "All Permissions") {
                if ($permission -ne $permFilter) { $include = $false }
            }

            if ($include) {
                $filtered += [PSCustomObject]@{
                    Name        = $name
                    Email       = $email
                    UserType    = $userType
                    Permission  = $permission
                    Site        = $site
                    IsSiteAdmin = $isSiteAdmin
                }
            }
        }

        $Controls.dgUsers.ItemsSource = $filtered
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Apply-UsersFilter"
    }
}

function Refresh-UsersDeepDiveData {
    <#
    .SYNOPSIS
    Refreshes user data from SharePoint
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

        $Controls.txtStatus.Text = "Refreshing user data..."
        Load-UsersDeepDiveData -Controls $Controls

        [System.Windows.MessageBox]::Show(
            "User data refreshed successfully!",
            "Success",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Information
        )
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Refresh-UsersDeepDiveData"
        $Controls.txtStatus.Text = "Error refreshing data"
    }
}

function Export-UsersDeepDiveData {
    <#
    .SYNOPSIS
    Exports user deep dive data to CSV
    #>
    param($Controls)

    try {
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "SharePoint_Users_DeepDive_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        if ($saveDialog.ShowDialog() -eq $true) {
            $users = Get-SharePointData -DataType "Users"

            $exportData = @()
            foreach ($user in $users) {
                $exportData += [PSCustomObject]@{
                    Name       = $user["Name"]
                    Email      = $user["Email"]
                    Permission = $user["Permission"]
                    Site       = $user["Site"]
                    IsExternal = $user["IsExternal"]
                    IsSiteAdmin = $user["IsSiteAdmin"]
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
        Write-ErrorLog -Message $_.Exception.Message -Location "Export-UsersDeepDiveData"
        [System.Windows.MessageBox]::Show(
            "Failed to export data: $($_.Exception.Message)",
            "Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}
