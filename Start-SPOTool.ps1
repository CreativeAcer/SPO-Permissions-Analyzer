#Requires -Version 7.0
<#
.SYNOPSIS
    SharePoint Online Permissions Report Tool
.DESCRIPTION
    Simple tool for SharePoint Online permissions analysis with persistent authentication
#>

Add-Type -AssemblyName PresentationFramework
Add-Type -AssemblyName System.Windows.Forms

# Import functions
. "$PSScriptRoot\Functions\Core\Settings.ps1"
. "$PSScriptRoot\Functions\Core\Logging.ps1"
. "$PSScriptRoot\Functions\Core\SharePointDataManager.ps1"
. "$PSScriptRoot\Functions\Core\ThrottleProtection.ps1"
. "$PSScriptRoot\Functions\Core\Checkpoint.ps1"
. "$PSScriptRoot\Functions\Core\JsonExport.ps1"
. "$PSScriptRoot\Functions\Core\GraphEnrichment.ps1"
. "$PSScriptRoot\Functions\Core\RiskScoring.ps1"
. "$PSScriptRoot\Functions\Core\AuditLog.ps1"
. "$PSScriptRoot\Functions\UI\UIManager.ps1"
. "$PSScriptRoot\Functions\SharePoint\SPOConnection.ps1"
. "$PSScriptRoot\Functions\UI\ConnectionTab.ps1"
. "$PSScriptRoot\Functions\UI\OperationsTab.ps1"
. "$PSScriptRoot\Functions\UI\VisualAnalyticsTab.ps1"
. "$PSScriptRoot\Functions\UI\HelpTab.ps1"
. "$PSScriptRoot\Functions\UI\MainWindow.ps1"

# DeepDive
. "$PSScriptRoot\Functions\UI\DeepDive\SitesDeepDive.ps1"
. "$PSScriptRoot\Functions\UI\DeepDive\UsersDeepDive.ps1"
. "$PSScriptRoot\Functions\UI\DeepDive\GroupsDeepDive.ps1"
. "$PSScriptRoot\Functions\UI\DeepDive\ExternalUsersDeepDive.ps1"
. "$PSScriptRoot\Functions\UI\DeepDive\PermissionsDeepDive.ps1"
. "$PSScriptRoot\Functions\UI\DeepDive\InheritanceDeepDive.ps1"
. "$PSScriptRoot\Functions\UI\DeepDive\SharingLinksDeepDive.ps1"

# Global variables
$script:SPOConnected = $false
$script:SPOContext = $null

# Initialize settings
Initialize-Settings

try {
    Write-Host "Starting SharePoint Permissions Tool..." -ForegroundColor Green
    
    # Show main window
    Show-MainWindow
}
catch {
    Write-ErrorLog -Message $_.Exception.Message -Location "Main"
    [System.Windows.MessageBox]::Show(
        "Failed to start application: $($_.Exception.Message)", 
        "Error", 
        [System.Windows.MessageBoxButton]::OK, 
        [System.Windows.MessageBoxImage]::Error
    )
}
finally {
    # Cleanup
    if ($script:SPOConnected) {
        try {
            Disconnect-PnPOnline -ErrorAction SilentlyContinue
        } catch {
            # Ignore cleanup errors
        }
    }
}
