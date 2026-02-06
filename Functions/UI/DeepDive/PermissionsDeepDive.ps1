# ============================================
# PermissionsDeepDive.ps1 - Role Assignment Mapping
# ============================================
# Location: Functions/UI/DeepDive/PermissionsDeepDive.ps1

function Show-PermissionsDeepDive {
    <#
    .SYNOPSIS
    Shows the Role Assignment Mapping deep dive window
    #>
    try {
        Write-ActivityLog "Opening Permissions Deep Dive window" -Level "Information"

        # Load XAML
        $xamlPath = Join-Path $PSScriptRoot "..\..\..\Views\DeepDive\PermissionsDeepDive.xaml"
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
            txtRoleCount = $deepDiveWindow.FindName("txtRoleCount")
            btnRefreshData = $deepDiveWindow.FindName("btnRefreshData")
            btnExport = $deepDiveWindow.FindName("btnExport")
            # Summary Stats
            txtTotalAssignments = $deepDiveWindow.FindName("txtTotalAssignments")
            txtFullControl = $deepDiveWindow.FindName("txtFullControl")
            txtEditCount = $deepDiveWindow.FindName("txtEditCount")
            txtReadOnly = $deepDiveWindow.FindName("txtReadOnly")
            txtCustomRoles = $deepDiveWindow.FindName("txtCustomRoles")
            # Role Assignments Tab
            txtSearch = $deepDiveWindow.FindName("txtSearch")
            cboRoleFilter = $deepDiveWindow.FindName("cboRoleFilter")
            cboPrincipalFilter = $deepDiveWindow.FindName("cboPrincipalFilter")
            dgRoleAssignments = $deepDiveWindow.FindName("dgRoleAssignments")
            # Role Distribution Tab
            canvasRoleDistribution = $deepDiveWindow.FindName("canvasRoleDistribution")
            dgRoleBreakdown = $deepDiveWindow.FindName("dgRoleBreakdown")
            # Security Review Tab
            txtHighRisk = $deepDiveWindow.FindName("txtHighRisk")
            txtMediumRisk = $deepDiveWindow.FindName("txtMediumRisk")
            txtLowRisk = $deepDiveWindow.FindName("txtLowRisk")
            lstSecurityFindings = $deepDiveWindow.FindName("lstSecurityFindings")
            # Status Bar
            txtStatus = $deepDiveWindow.FindName("txtStatus")
            txtLastUpdate = $deepDiveWindow.FindName("txtLastUpdate")
        }

        # Set up event handlers
        $controls.btnRefreshData.Add_Click({
            Refresh-PermissionsDeepDiveData -Controls $controls
        })

        $controls.btnExport.Add_Click({
            Export-PermissionsDeepDiveData -Controls $controls
        })

        $controls.txtSearch.Add_TextChanged({
            Apply-PermissionsFilter -Controls $controls
        })

        $controls.cboRoleFilter.Add_SelectionChanged({
            Apply-PermissionsFilter -Controls $controls
        })

        $controls.cboPrincipalFilter.Add_SelectionChanged({
            Apply-PermissionsFilter -Controls $controls
        })

        # Load initial data
        Load-PermissionsDeepDiveData -Controls $controls

        # Show window
        $deepDiveWindow.ShowDialog() | Out-Null

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Show-PermissionsDeepDive"
        [System.Windows.MessageBox]::Show(
            "Failed to open Permissions Deep Dive: $($_.Exception.Message)",
            "Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}

function Load-PermissionsDeepDiveData {
    <#
    .SYNOPSIS
    Loads data into the Permissions Deep Dive window
    #>
    param($Controls)

    try {
        $Controls.txtStatus.Text = "Loading role assignment data..."

        $roleAssignments = Get-SharePointData -DataType "RoleAssignments"

        # Update header
        $Controls.txtRoleCount.Text = "Analyzing $($roleAssignments.Count) role assignments"

        # Calculate summary statistics
        $fullControlCount = 0
        $editCount = 0
        $readOnlyCount = 0
        $customCount = 0

        foreach ($ra in $roleAssignments) {
            $role = if ($ra["Role"]) { $ra["Role"] } else { "Unknown" }
            switch -Wildcard ($role) {
                "Full Control" { $fullControlCount++ }
                "Edit" { $editCount++ }
                "Contribute" { $editCount++ }
                "Read" { $readOnlyCount++ }
                "View Only" { $readOnlyCount++ }
                default { $customCount++ }
            }
        }

        $Controls.txtTotalAssignments.Text = $roleAssignments.Count.ToString()
        $Controls.txtFullControl.Text = $fullControlCount.ToString()
        $Controls.txtEditCount.Text = $editCount.ToString()
        $Controls.txtReadOnly.Text = $readOnlyCount.ToString()
        $Controls.txtCustomRoles.Text = $customCount.ToString()

        # Load grid
        Load-RoleAssignmentsGrid -Controls $Controls -Roles $roleAssignments

        # Load distribution chart
        Load-RoleDistribution -Controls $Controls -Roles $roleAssignments

        # Load security review
        Load-PermissionsSecurityReview -Controls $Controls -Roles $roleAssignments

        $Controls.txtStatus.Text = "Ready"
        $Controls.txtLastUpdate.Text = "Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-PermissionsDeepDiveData"
        $Controls.txtStatus.Text = "Error loading data"
    }
}

function Load-RoleAssignmentsGrid {
    <#
    .SYNOPSIS
    Loads the role assignments data grid
    #>
    param($Controls, $Roles)

    try {
        $roleObjects = @()

        foreach ($ra in $Roles) {
            $roleObjects += [PSCustomObject]@{
                Principal = if ($ra["Principal"]) { $ra["Principal"] } else { "Unknown" }
                PrincipalType = if ($ra["PrincipalType"]) { $ra["PrincipalType"] } else { "Unknown" }
                Role = if ($ra["Role"]) { $ra["Role"] } else { "Unknown" }
                Scope = if ($ra["Scope"]) { $ra["Scope"] } else { "Site" }
                ScopeUrl = if ($ra["ScopeUrl"]) { $ra["ScopeUrl"] } else { "N/A" }
                SiteTitle = if ($ra["SiteTitle"]) { $ra["SiteTitle"] } else { "N/A" }
            }
        }

        $Controls.dgRoleAssignments.ItemsSource = $roleObjects

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-RoleAssignmentsGrid"
    }
}

function Load-RoleDistribution {
    <#
    .SYNOPSIS
    Loads role distribution chart and breakdown
    #>
    param($Controls, $Roles)

    try {
        $Controls.canvasRoleDistribution.Children.Clear()

        # Count by role
        $roleCounts = @{}
        foreach ($ra in $Roles) {
            $role = if ($ra["Role"]) { $ra["Role"] } else { "Unknown" }
            if ($roleCounts.ContainsKey($role)) {
                $roleCounts[$role]++
            } else {
                $roleCounts[$role] = 1
            }
        }

        if ($roleCounts.Count -eq 0) { return }

        $total = ($roleCounts.Values | Measure-Object -Sum).Sum
        if ($total -eq 0) { $total = 1 }

        $colorMap = @{
            "Full Control" = "#DC3545"
            "Edit"         = "#FFC107"
            "Contribute"   = "#FD7E14"
            "Read"         = "#28A745"
            "View Only"    = "#17A2B8"
            "Limited Access" = "#6C757D"
        }

        # Build breakdown data
        $breakdownData = @()
        $rank = 1
        foreach ($key in $roleCounts.Keys | Sort-Object { $roleCounts[$_] } -Descending) {
            $count = $roleCounts[$key]
            $percentage = [math]::Round(($count / $total) * 100, 1)
            $color = if ($colorMap.ContainsKey($key)) { $colorMap[$key] } else { "#6F42C1" }

            $breakdownData += [PSCustomObject]@{
                Role = $key
                Count = $count
                Percentage = "$percentage%"
                PercentageValue = $percentage
                BarColor = $color
            }
            $rank++
        }

        $Controls.dgRoleBreakdown.ItemsSource = $breakdownData

        # Draw bar chart
        $startY = 20
        $barHeight = 25
        $maxBarWidth = 150

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
            [System.Windows.Controls.Canvas]::SetTop($bar, $startY + ($i * ($barHeight + 10)))
            $Controls.canvasRoleDistribution.Children.Add($bar)

            $label = New-Object System.Windows.Controls.TextBlock
            $label.Text = "$($item.Role): $($item.Count) ($($item.Percentage))"
            $label.FontSize = 10
            $label.Foreground = "#495057"

            [System.Windows.Controls.Canvas]::SetLeft($label, $barWidth + 20)
            [System.Windows.Controls.Canvas]::SetTop($label, $startY + ($i * ($barHeight + 10)) + 5)
            $Controls.canvasRoleDistribution.Children.Add($label)
        }

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-RoleDistribution"
    }
}

function Load-PermissionsSecurityReview {
    <#
    .SYNOPSIS
    Loads security review findings for role assignments
    #>
    param($Controls, $Roles)

    try {
        $highRisk = 0
        $mediumRisk = 0
        $lowRisk = 0
        $findings = @()

        # Count Full Control assignments (high risk if many)
        $fullControlAssignments = $Roles | Where-Object { $_["Role"] -eq "Full Control" }
        if ($fullControlAssignments.Count -gt 5) {
            $highRisk++
            $findings += [PSCustomObject]@{
                Icon = "🚨"
                Site = "Tenant-wide"
                Issue = "$($fullControlAssignments.Count) principals have Full Control access"
                Recommendation = "Review and reduce Full Control assignments to minimum necessary"
                SeverityColor = "#FFCDD2"
            }
        } elseif ($fullControlAssignments.Count -gt 2) {
            $mediumRisk++
            $findings += [PSCustomObject]@{
                Icon = "⚠️"
                Site = "Tenant-wide"
                Issue = "$($fullControlAssignments.Count) principals have Full Control access"
                Recommendation = "Verify each Full Control assignment is justified"
                SeverityColor = "#FFE0B2"
            }
        } else {
            $lowRisk++
        }

        # Check for external users with edit+ permissions
        $externalEditors = $Roles | Where-Object {
            $_["PrincipalType"] -eq "External" -and
            $_["Role"] -in @("Full Control", "Edit", "Contribute")
        }
        if ($externalEditors.Count -gt 0) {
            $highRisk++
            $findings += [PSCustomObject]@{
                Icon = "🚨"
                Site = "Multiple Sites"
                Issue = "$($externalEditors.Count) external users have edit or higher permissions"
                Recommendation = "Audit external user access and restrict to read-only where possible"
                SeverityColor = "#FFCDD2"
            }
        }

        # Check for groups with broad access
        $groupFullControl = $Roles | Where-Object {
            $_["PrincipalType"] -eq "SharePoint Group" -and $_["Role"] -eq "Full Control"
        }
        if ($groupFullControl.Count -gt 3) {
            $mediumRisk++
            $findings += [PSCustomObject]@{
                Icon = "⚠️"
                Site = "Multiple Sites"
                Issue = "$($groupFullControl.Count) groups have Full Control"
                Recommendation = "Consider using more granular permission levels for groups"
                SeverityColor = "#FFE0B2"
            }
        }

        # Check for scope spread (assignments at list/library level)
        $listLevelAssignments = $Roles | Where-Object { $_["Scope"] -in @("List", "Library") }
        if ($listLevelAssignments.Count -gt 10) {
            $mediumRisk++
            $findings += [PSCustomObject]@{
                Icon = "⚠️"
                Site = "Multiple locations"
                Issue = "$($listLevelAssignments.Count) item-level permission assignments found"
                Recommendation = "Consider simplifying by using site-level permissions with inheritance"
                SeverityColor = "#FFE0B2"
            }
        }

        # Positive findings
        if ($findings.Count -eq 0) {
            $findings += [PSCustomObject]@{
                Icon = "✅"
                Site = "Tenant-wide"
                Issue = "No significant permission risks detected"
                Recommendation = "Continue monitoring permissions regularly"
                SeverityColor = "#C8E6C9"
            }
        }

        $findings += [PSCustomObject]@{
            Icon = "📊"
            Site = "Summary"
            Issue = "Analyzed $($Roles.Count) role assignments across all scopes"
            Recommendation = "Schedule regular permission reviews to maintain security posture"
            SeverityColor = "#E1F5FE"
        }

        $Controls.txtHighRisk.Text = $highRisk.ToString()
        $Controls.txtMediumRisk.Text = $mediumRisk.ToString()
        $Controls.txtLowRisk.Text = $lowRisk.ToString()
        $Controls.lstSecurityFindings.ItemsSource = $findings

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Load-PermissionsSecurityReview"
    }
}

function Apply-PermissionsFilter {
    <#
    .SYNOPSIS
    Applies filters to the role assignments data grid
    #>
    param($Controls)

    try {
        $searchText = $Controls.txtSearch.Text.ToLower()
        $roleFilter = $Controls.cboRoleFilter.SelectedItem.Content
        $principalFilter = $Controls.cboPrincipalFilter.SelectedItem.Content

        $roles = Get-SharePointData -DataType "RoleAssignments"
        $filtered = @()

        foreach ($ra in $roles) {
            $include = $true

            # Search filter
            if ($searchText -and $searchText.Length -gt 0) {
                $principal = if ($ra["Principal"]) { $ra["Principal"].ToLower() } else { "" }
                $role = if ($ra["Role"]) { $ra["Role"].ToLower() } else { "" }
                $scope = if ($ra["ScopeUrl"]) { $ra["ScopeUrl"].ToLower() } else { "" }

                if (-not ($principal.Contains($searchText) -or $role.Contains($searchText) -or $scope.Contains($searchText))) {
                    $include = $false
                }
            }

            # Role filter
            if ($roleFilter -ne "All Roles" -and $include) {
                $raRole = if ($ra["Role"]) { $ra["Role"] } else { "" }
                if ($roleFilter -eq "Custom") {
                    $standardRoles = @("Full Control", "Edit", "Contribute", "Read", "View Only")
                    if ($raRole -in $standardRoles) { $include = $false }
                } elseif ($raRole -ne $roleFilter) {
                    $include = $false
                }
            }

            # Principal type filter
            if ($principalFilter -ne "All Principals" -and $include) {
                $pType = if ($ra["PrincipalType"]) { $ra["PrincipalType"] } else { "" }
                switch ($principalFilter) {
                    "Users" { if ($pType -ne "User" -and $pType -ne "External") { $include = $false } }
                    "Groups" { if ($pType -notlike "*Group*") { $include = $false } }
                }
            }

            if ($include) { $filtered += $ra }
        }

        Load-RoleAssignmentsGrid -Controls $Controls -Roles $filtered

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Apply-PermissionsFilter"
    }
}

function Refresh-PermissionsDeepDiveData {
    <#
    .SYNOPSIS
    Refreshes the permissions deep dive data
    #>
    param($Controls)

    try {
        $Controls.txtStatus.Text = "Refreshing data..."

        if ($script:SPOConnected) {
            Load-PermissionsDeepDiveData -Controls $Controls

            [System.Windows.MessageBox]::Show(
                "Data refreshed successfully!",
                "Success",
                [System.Windows.MessageBoxButton]::OK,
                [System.Windows.MessageBoxImage]::Information
            )
        }
        else {
            # Reload from existing data store
            Load-PermissionsDeepDiveData -Controls $Controls
        }

    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Refresh-PermissionsDeepDiveData"
        $Controls.txtStatus.Text = "Error refreshing data"
    }
}

function Export-PermissionsDeepDiveData {
    <#
    .SYNOPSIS
    Exports the role assignment data to CSV
    #>
    param($Controls)

    try {
        $saveDialog = New-Object Microsoft.Win32.SaveFileDialog
        $saveDialog.Filter = "CSV files (*.csv)|*.csv|All files (*.*)|*.*"
        $saveDialog.FileName = "SharePoint_RoleAssignments_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"

        if ($saveDialog.ShowDialog() -eq $true) {
            $roles = Get-SharePointData -DataType "RoleAssignments"

            $exportData = @()
            foreach ($ra in $roles) {
                $exportData += [PSCustomObject]@{
                    Principal = $ra["Principal"]
                    PrincipalType = $ra["PrincipalType"]
                    Role = $ra["Role"]
                    Scope = $ra["Scope"]
                    ScopeUrl = $ra["ScopeUrl"]
                    SiteTitle = $ra["SiteTitle"]
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
        Write-ErrorLog -Message $_.Exception.Message -Location "Export-PermissionsDeepDiveData"
        [System.Windows.MessageBox]::Show(
            "Failed to export data: $($_.Exception.Message)",
            "Export Error",
            [System.Windows.MessageBoxButton]::OK,
            [System.Windows.MessageBoxImage]::Error
        )
    }
}
