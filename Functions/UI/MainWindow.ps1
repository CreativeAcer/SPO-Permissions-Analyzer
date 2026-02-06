function Show-MainWindow {
    <#
    .SYNOPSIS
    Main window coordinator that loads XAML and initializes all tab modules
    #>
    try {
        Write-ActivityLog "Loading MainWindow XAML from external file" -Level "Information"
        
        # Get the project root directory
        $projectRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
        $xamlPath = Join-Path $projectRoot "Views\Windows\MainWindow.xaml"
        
        # Check if XAML file exists
        if (-not (Test-Path $xamlPath)) {
            throw "XAML file not found at: $xamlPath. Please ensure the Views\Windows\MainWindow.xaml file exists."
        }
        
        # Load XAML content
        try {
            $xamlContent = Get-Content $xamlPath -Raw -Encoding UTF8
            Write-ActivityLog "Successfully loaded XAML content from: $xamlPath" -Level "Information"
            
            # Load required assemblies
            Add-Type -AssemblyName System.Windows.Forms
            Add-Type -AssemblyName PresentationFramework
            Add-Type -AssemblyName PresentationCore
        }
        catch {
            throw "Failed to read XAML file: $($_.Exception.Message)"
        }
        
        # Parse XAML
        try {
            $reader = [System.Xml.XmlNodeReader]::new([xml]$xamlContent)
            $script:MainWindow = [System.Windows.Markup.XamlReader]::Load($reader)
            Write-ActivityLog "Successfully parsed XAML and created window" -Level "Information"
        }
        catch {
            throw "Failed to parse XAML content: $($_.Exception.Message)"
        }
        
        # Initialize all UI controls as global script variables
        Initialize-UIControls
        
        # Load saved settings
        Load-SavedSettings
        
        # Initialize all tab modules
        Initialize-AllTabs
        
        # Show window
        Write-ActivityLog "Displaying main window" -Level "Information"
        $script:MainWindow.ShowDialog() | Out-Null
        
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Show-MainWindow"
        throw
    }
}

function Initialize-UIControls {
    <#
    .SYNOPSIS
    Initializes all UI controls and stores them as script variables
    #>
    try {
        # Connection Tab Controls
        $script:txtTenantUrl = $script:MainWindow.FindName("txtTenantUrl")
        $script:txtClientId = $script:MainWindow.FindName("txtClientId")
        $script:btnConnectSPO = $script:MainWindow.FindName("btnConnectSPO")
        $script:btnDemoMode = $script:MainWindow.FindName("btnDemoMode")
        $script:txtSPOResults = $script:MainWindow.FindName("txtSPOResults")
        
        # Operations Tab Controls
        $script:txtSiteUrl = $script:MainWindow.FindName("txtSiteUrl")
        $script:btnGetSites = $script:MainWindow.FindName("btnGetSites")
        $script:btnGetPermissions = $script:MainWindow.FindName("btnGetPermissions")
        $script:btnGenerateReport = $script:MainWindow.FindName("btnGenerateReport")
        $script:txtStatus = $script:MainWindow.FindName("txtStatus")
        $script:txtOperationsResults = $script:MainWindow.FindName("txtOperationsResults")
        
        # Visual Analytics Tab Controls
        $script:txtAnalyticsTitle = $script:MainWindow.FindName("txtAnalyticsTitle")
        $script:txtAnalyticsSubtitle = $script:MainWindow.FindName("txtAnalyticsSubtitle")
        $script:btnRefreshAnalytics = $script:MainWindow.FindName("btnRefreshAnalytics")
        $script:txtTotalSites = $script:MainWindow.FindName("txtTotalSites")
        $script:txtTotalUsers = $script:MainWindow.FindName("txtTotalUsers")
        $script:txtTotalGroups = $script:MainWindow.FindName("txtTotalGroups")
        $script:txtExternalUsers = $script:MainWindow.FindName("txtExternalUsers")
        $script:canvasStorageChart = $script:MainWindow.FindName("canvasStorageChart")
        $script:canvasPermissionChart = $script:MainWindow.FindName("canvasPermissionChart")
        $script:dgSites = $script:MainWindow.FindName("dgSites")
        $script:lstPermissionAlerts = $script:MainWindow.FindName("lstPermissionAlerts")

        # Security Analysis Cards (P3)
        $script:txtRoleAssignments = $script:MainWindow.FindName("txtRoleAssignments")
        $script:txtInheritanceBreaks = $script:MainWindow.FindName("txtInheritanceBreaks")
        $script:txtSharingLinks = $script:MainWindow.FindName("txtSharingLinks")
        
        # Validate critical controls
        $criticalControls = @(
            @{Name="txtTenantUrl"; Control=$script:txtTenantUrl},
            @{Name="txtClientId"; Control=$script:txtClientId},
            @{Name="btnConnectSPO"; Control=$script:btnConnectSPO},
            @{Name="txtSPOResults"; Control=$script:txtSPOResults},
            @{Name="txtOperationsResults"; Control=$script:txtOperationsResults}
        )
        
        foreach ($controlCheck in $criticalControls) {
            if ($null -eq $controlCheck.Control) {
                throw "Critical control '$($controlCheck.Name)' not found in XAML. Please check the control names in MainWindow.xaml"
            }
        }
        
        Write-ActivityLog "All UI controls successfully loaded and validated" -Level "Information"
    }
    catch {
        throw "Failed to initialize UI controls: $($_.Exception.Message)"
    }
}

function Load-SavedSettings {
    <#
    .SYNOPSIS
    Loads previously saved application settings
    #>
    try {
        $savedTenantUrl = Get-AppSetting -SettingName "SharePoint.TenantUrl"
        if ($savedTenantUrl) {
            $script:txtTenantUrl.Text = $savedTenantUrl
            Write-ActivityLog "Loaded saved tenant URL" -Level "Information"
        }
        
        $savedClientId = Get-AppSetting -SettingName "SharePoint.ClientId"
        if ($savedClientId) {
            $script:txtClientId.Text = $savedClientId
            Write-ActivityLog "Loaded saved client ID" -Level "Information"
        }
    }
    catch {
        Write-ActivityLog "Warning: Could not load saved settings: $($_.Exception.Message)" -Level "Warning"
    }
}

function Initialize-AllTabs {
    <#
    .SYNOPSIS
    Initializes all tab modules in the correct order
    #>
    try {
        Write-ActivityLog "Initializing all tab modules" -Level "Information"
        
        # Initialize each tab module
        Initialize-ConnectionTab
        Initialize-OperationsTab
        Initialize-VisualAnalyticsTab
        
        Write-ActivityLog "All tab modules initialized successfully" -Level "Information"
    }
    catch {
        throw "Failed to initialize tab modules: $($_.Exception.Message)"
    }
}

# Global helper function for UI updates during long operations
function Update-UIAndWait {
    <#
    .SYNOPSIS
    Forces UI update and optional wait - helps with responsiveness
    #>
    param(
        [int]$WaitMs = 0
    )
    
    try {
        [System.Windows.Forms.Application]::DoEvents()
        [System.Windows.Threading.Dispatcher]::CurrentDispatcher.Invoke([System.Windows.Threading.DispatcherPriority]::Background, [action]{})
        
        if ($WaitMs -gt 0) {
            Start-Sleep -Milliseconds $WaitMs
        }
    }
    catch {
        # Silently continue if UI update fails
    }
}

# Global helper function for updating console with real-time output
function Write-ConsoleOutput {
    <#
    .SYNOPSIS
    Writes output to the operations console with real-time updates
    #>
    param(
        [string]$Message,
        [switch]$Append,
        [switch]$NewLine = $true,
        [switch]$ForceUpdate
    )
    
    try {
        if ($Append) {
            if ($NewLine) {
                $script:txtOperationsResults.Text += "$Message`n"
            } else {
                $script:txtOperationsResults.Text += $Message
            }
        } else {
            $script:txtOperationsResults.Text = if ($NewLine) { "$Message`n" } else { $Message }
        }
        
        # Auto-scroll to bottom
        $script:txtOperationsResults.ScrollToEnd()
        
        # Force UI update for real-time display
        if ($ForceUpdate) {
            Update-UIAndWait -WaitMs 50
        }
    }
    catch {
        # Fallback to Write-Host if UI update fails
        Write-Host $Message
    }
}