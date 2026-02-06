#Requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Online Permissions Analyzer - Web UI Version
.DESCRIPTION
    Launches a local web server that serves a browser-based UI.
    The SharePoint/Core backend is identical to the WPF version;
    only the presentation layer changes (HTML/CSS/JS instead of XAML/WPF).
.NOTES
    Run with: pwsh ./Start-SPOTool-Web.ps1
    Opens http://localhost:8080 in your default browser.
#>

param(
    [int]$Port = 8080
)

# ============================================
# 1. Load Core modules (shared with WPF version)
# ============================================
. "$PSScriptRoot\Functions\Core\Logging.ps1"
. "$PSScriptRoot\Functions\Core\Settings.ps1"
. "$PSScriptRoot\Functions\Core\SharePointDataManager.ps1"

# ============================================
# 2. Load SharePoint modules (shared with WPF version)
# ============================================
. "$PSScriptRoot\Functions\SharePoint\SPOConnection.ps1"

# ============================================
# 3. Load Operations (data collection logic, reused as-is)
# ============================================
. "$PSScriptRoot\Functions\UI\OperationsTab.ps1"

# ============================================
# 4. Load Web Server modules (replaces WPF UI layer)
# ============================================
. "$PSScriptRoot\Functions\Server\WebServer.ps1"
. "$PSScriptRoot\Functions\Server\ApiHandlers.ps1"

# ============================================
# 5. Initialize and start
# ============================================
Write-ActivityLog "=== SharePoint Permissions Analyzer (Web UI) ===" -Level "Information"

# Override Write-ConsoleOutput for web mode - redirect to operation log
function Write-ConsoleOutput {
    param(
        [string]$Message,
        [switch]$Append,
        [switch]$NewLine = $true,
        [switch]$ForceUpdate
    )
    Add-OperationLog -Message $Message
}

# Override Update-UIAndWait for web mode - no-op
function Update-UIAndWait {
    param([int]$WaitMs = 0)
    # No WPF dispatcher needed in web mode
}

# Initialize globals
$script:SPOConnected = $false
$script:DemoMode = $false

Initialize-SharePointDataManager

# Start the web server (blocks until Ctrl+C)
Start-WebServer -Port $Port
