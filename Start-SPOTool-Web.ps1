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
    [int]$Port = 8080,
    [string]$ListenAddress = "localhost",
    [switch]$NoBrowser
)

# ============================================
# 1. Load Core modules (shared with WPF version)
# ============================================
. "$PSScriptRoot\Functions\Core\Logging.ps1"
. "$PSScriptRoot\Functions\Core\Settings.ps1"
. "$PSScriptRoot\Functions\Core\SharePointDataManager.ps1"
. "$PSScriptRoot\Functions\Core\ThrottleProtection.ps1"
. "$PSScriptRoot\Functions\Core\Checkpoint.ps1"
. "$PSScriptRoot\Functions\Core\JsonExport.ps1"
. "$PSScriptRoot\Functions\Core\GraphEnrichment.ps1"
. "$PSScriptRoot\Functions\Core\RiskScoring.ps1"
. "$PSScriptRoot\Functions\Core\AuditLog.ps1"

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

# ============================================
# 6. Auto-connect if env vars are set (container mode)
# ============================================
if ($env:SPO_TENANT_URL -and $env:SPO_CLIENT_ID) {
    Write-Host ""
    Write-Host "  SharePoint credentials detected in environment." -ForegroundColor Cyan
    Write-Host "  Tenant: $($env:SPO_TENANT_URL)" -ForegroundColor White
    Write-Host "  Client: $($env:SPO_CLIENT_ID)" -ForegroundColor White
    Write-Host ""

    try {
        if (-not (Test-PnPModuleAvailable)) {
            throw "PnP.PowerShell module not available"
        }

        Write-Host "  Connecting via device code flow..." -ForegroundColor Yellow
        Write-Host "  Follow the instructions below to authenticate:" -ForegroundColor Yellow
        Write-Host ""

        Connect-PnPOnline -Url $env:SPO_TENANT_URL -ClientId $env:SPO_CLIENT_ID -DeviceLogin

        $script:SPOConnected = $true
        Set-AppSetting -SettingName "SharePoint.TenantUrl" -Value $env:SPO_TENANT_URL
        Set-AppSetting -SettingName "SharePoint.ClientId" -Value $env:SPO_CLIENT_ID

        Write-Host ""
        Write-Host "  Connected to SharePoint Online!" -ForegroundColor Green
        Write-Host ""
    }
    catch {
        Write-Host ""
        Write-Host "  Auto-connect failed: $($_.Exception.Message)" -ForegroundColor Red
        Write-Host "  Starting in disconnected mode. Use Demo Mode or connect via the UI." -ForegroundColor Yellow
        Write-Host ""
    }
}

# Start the web server (blocks until Ctrl+C)
Start-WebServer -Port $Port -ListenAddress $ListenAddress -NoBrowser:$NoBrowser
