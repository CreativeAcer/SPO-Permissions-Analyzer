# ============================================
# BackgroundJobManager.ps1 - Background Operation Orchestration
# ============================================
# Manages background runspaces for long-running operations
# (site enumeration, permissions analysis, enrichment).
# Keeps the HTTP server responsive while operations run.

function Start-BackgroundOperation {
    <#
    .SYNOPSIS
    Runs a long-running operation in a background runspace so the HTTP server stays responsive.
    The scriptblock receives $SharedState (synchronized hashtable) and $ScriptRoot (project root).
    All core modules are dot-sourced into the new runspace and the PnP connection is re-established
    via access token forwarding.
    #>
    param(
        [Parameter(Mandatory)]
        [scriptblock]$ScriptBlock
    )

    # Capture the PnP access token from the current session so the background runspace
    # can re-connect without an interactive prompt
    $accessToken = $null
    $tenantUrl = $null
    $clientId = $null
    try {
        $accessToken = Get-PnPAccessToken -ErrorAction SilentlyContinue
        $tenantUrl = (Get-AppSetting -SettingName "SharePoint.TenantUrl")
        $clientId = (Get-AppSetting -SettingName "SharePoint.ClientId")
    } catch { }

    $projectRoot = (Resolve-Path (Join-Path $PSScriptRoot "..\..")).Path

    # Prepare initial session state with required variables
    $iss = [System.Management.Automation.Runspaces.InitialSessionState]::CreateDefault()

    $ps = [PowerShell]::Create()
    $ps.Runspace = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspace($iss)
    $ps.Runspace.Open()

    # Pass shared state and project root into the runspace
    $ps.Runspace.SessionStateProxy.SetVariable('SharedState', $script:ServerState)
    $ps.Runspace.SessionStateProxy.SetVariable('ScriptRoot', $projectRoot)
    $ps.Runspace.SessionStateProxy.SetVariable('AccessToken', $accessToken)
    $ps.Runspace.SessionStateProxy.SetVariable('TenantUrl', $tenantUrl)
    $ps.Runspace.SessionStateProxy.SetVariable('ClientId', $clientId)

    # The wrapper script loads all modules, re-establishes PnP connection, then runs the operation
    $wrapperScript = {
        param($OperationScript)

        # Dot-source all core, analysis, SharePoint, and demo modules
        . "$ScriptRoot\Functions\Core\Logging.ps1"
        . "$ScriptRoot\Functions\Core\Settings.ps1"
        . "$ScriptRoot\Functions\Core\OutputAdapter.ps1"
        . "$ScriptRoot\Functions\Core\SharePointDataManager.ps1"
        . "$ScriptRoot\Functions\Core\ThrottleProtection.ps1"
        . "$ScriptRoot\Functions\Core\Checkpoint.ps1"
        . "$ScriptRoot\Functions\Core\AuditLog.ps1"
        . "$ScriptRoot\Functions\Analysis\JsonExport.ps1"
        . "$ScriptRoot\Functions\Analysis\GraphEnrichment.ps1"
        . "$ScriptRoot\Functions\Analysis\RiskScoring.ps1"
        . "$ScriptRoot\Functions\SharePoint\SPOConnection.ps1"
        . "$ScriptRoot\Functions\SharePoint\SiteCollector.ps1"
        . "$ScriptRoot\Functions\SharePoint\PermissionsCollector.ps1"
        . "$ScriptRoot\Functions\Demo\DemoDataGenerator.ps1"

        # Override Write-ConsoleOutput to write to the shared operation log
        function Write-ConsoleOutput {
            param(
                [string]$Message,
                [switch]$Append,
                [switch]$NewLine = $true,
                [switch]$ForceUpdate
            )
            [void]$SharedState.OperationLog.Add($Message)
        }

        # Override Update-UIAndWait - no-op in background
        function Update-UIAndWait {
            param([int]$WaitMs = 0)
        }

        # Point the data manager at the SAME synchronized data store from the main runspace
        # This is safe because SharedState is a synchronized hashtable
        $script:SharePointData = $SharedState.SharePointData

        # Re-establish PnP connection in this runspace.
        # Use the operation-specific site URL if available (e.g. for permissions analysis),
        # otherwise fall back to the tenant root URL (e.g. for site enumeration).
        $connectUrl = if ($SharedState.OperationSiteUrl) { $SharedState.OperationSiteUrl } else { $TenantUrl }
        if ($connectUrl) {
            $connected = $false

            # Preferred: use -ClientId -Interactive, which leverages MSAL's persistent
            # token cache from the user's earlier interactive login.  MSAL will silently
            # acquire a token without showing a browser popup.
            if ($ClientId) {
                try {
                    Connect-PnPOnline -Url $connectUrl -ClientId $ClientId -Interactive -ErrorAction Stop
                    $connected = $true
                } catch {
                    [void]$SharedState.OperationLog.Add("Note: Silent auth failed, trying access token fallback...")
                }
            }

            # Fallback: use the forwarded access token string
            if (-not $connected -and $AccessToken) {
                try {
                    Connect-PnPOnline -Url $connectUrl -AccessToken $AccessToken -ErrorAction Stop
                    $connected = $true
                } catch {
                    [void]$SharedState.OperationLog.Add("Warning: Could not re-establish PnP connection in background: $($_.Exception.Message)")
                }
            }
        }

        try {
            $SharedState.OperationRunning = $true
            $SharedState.OperationComplete = $false

            # Re-create the scriptblock in THIS runspace's session state so that
            # variables like $SharedState, $ScriptRoot, $AccessToken resolve here
            # instead of in the main thread's session (where they don't exist).
            $localScript = [scriptblock]::Create($OperationScript.ToString())
            & $localScript

            $SharedState.OperationRunning = $false
            $SharedState.OperationComplete = $true
        }
        catch {
            [void]$SharedState.OperationLog.Add("Error: $($_.Exception.Message)")
            $SharedState.OperationRunning = $false
            $SharedState.OperationComplete = $true
            $SharedState.OperationError = $_.Exception.Message
        }
    }

    [void]$ps.AddScript($wrapperScript).AddArgument($ScriptBlock)
    $script:ServerState.BackgroundJob = $ps.BeginInvoke()
}
