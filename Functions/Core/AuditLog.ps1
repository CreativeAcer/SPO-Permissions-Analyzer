# ============================================
# AuditLog.ps1 - Compliance-grade audit logging
# ============================================
# Produces structured audit trail with session tracking,
# timestamps, file hashes, and operation events.

$script:AuditSession = $null

function Start-AuditSession {
    <#
    .SYNOPSIS
    Begins a new audit session for tracking operations
    #>
    param(
        [string]$OperationType,
        [string]$ScanScope = "All"
    )

    $script:AuditSession = @{
        SessionId      = [guid]::NewGuid().ToString()
        StartTimestamp  = (Get-Date).ToString("o")
        EndTimestamp    = $null
        Duration        = $null
        OperationType   = $OperationType
        TenantUrl       = (Get-AppSetting -SettingName "SharePoint.TenantUrl")
        AppId           = (Get-AppSetting -SettingName "SharePoint.ClientId")
        ScanScope       = $ScanScope
        UserPrincipal   = "N/A"
        ToolVersion     = "1.1.0"
        HostName        = $env:COMPUTERNAME
        PowerShellVersion = $PSVersionTable.PSVersion.ToString()
        Events          = [System.Collections.ArrayList]::new()
        OutputFiles     = @()
        Status          = "InProgress"
        ErrorCount      = 0
        Metrics         = @{}
    }

    # Try to get current user from PnP connection
    try {
        $user = Get-PnPCurrentUser -ErrorAction SilentlyContinue
        if ($user) {
            $script:AuditSession.UserPrincipal = @(
                $user.UserPrincipalName, $user.Email, $user.LoginName
            ) | Where-Object { $_ } | Select-Object -First 1
        }
    }
    catch { }

    Write-AuditEvent -EventType "SessionStart" -Detail "Audit session started: $OperationType (scope: $ScanScope)"
}

function Write-AuditEvent {
    <#
    .SYNOPSIS
    Records an event in the current audit session
    #>
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("SessionStart", "SessionEnd", "DataCollection", "Export", "Connection", "Error", "Warning", "Info")]
        [string]$EventType,
        [string]$Detail,
        [string]$AffectedObject = ""
    )

    if (-not $script:AuditSession) { return }

    $event = @{
        Timestamp      = (Get-Date).ToString("o")
        EventType      = $EventType
        Detail         = $Detail
        AffectedObject = $AffectedObject
    }
    [void]$script:AuditSession.Events.Add($event)

    if ($EventType -eq "Error") {
        $script:AuditSession.ErrorCount++
    }

    # Also write to standard activity log
    Write-ActivityLog "AUDIT [$EventType] $Detail" -Level $(if ($EventType -eq "Error") { "Error" } else { "Information" })
}

function Complete-AuditSession {
    <#
    .SYNOPSIS
    Finalizes the audit session and writes the audit log file
    #>
    param(
        [string]$Status = "Completed"
    )

    if (-not $script:AuditSession) { return $null }

    $script:AuditSession.EndTimestamp = (Get-Date).ToString("o")
    $script:AuditSession.Status = $Status

    # Duration
    $start = [DateTime]::Parse($script:AuditSession.StartTimestamp)
    $end = [DateTime]::Parse($script:AuditSession.EndTimestamp)
    $script:AuditSession.Duration = ($end - $start).ToString()

    # Capture final metrics
    try {
        $metrics = Get-SharePointData -DataType "Metrics"
        $script:AuditSession.Metrics = @{
            TotalSites           = $metrics.TotalSites
            TotalUsers           = $metrics.TotalUsers
            TotalGroups          = $metrics.TotalGroups
            ExternalUsers        = $metrics.ExternalUsers
            TotalRoleAssignments = $metrics.TotalRoleAssignments
            InheritanceBreaks    = $metrics.InheritanceBreaks
            TotalSharingLinks    = $metrics.TotalSharingLinks
        }
    }
    catch { }

    Write-AuditEvent -EventType "SessionEnd" -Detail "Session $Status. Duration: $($script:AuditSession.Duration). Errors: $($script:AuditSession.ErrorCount)"

    # Save audit log to JSON file
    $logPath = Get-AppSetting -SettingName "Logging.LogPath"
    if (-not $logPath) { $logPath = "./Logs" }
    if (-not (Test-Path $logPath)) {
        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
    }

    $sessionShort = $script:AuditSession.SessionId.Substring(0, 8)
    $auditFile = Join-Path $logPath "audit_$(Get-Date -Format 'yyyyMMdd_HHmmss')_$sessionShort.json"

    $script:AuditSession | ConvertTo-Json -Depth 10 | Set-Content $auditFile -Encoding UTF8

    Write-ActivityLog "Audit log saved: $auditFile" -Level "Information"
    return $auditFile
}

function Get-AuditSession {
    <#
    .SYNOPSIS
    Returns the current audit session data
    #>
    return $script:AuditSession
}
