# ============================================
# Checkpoint.ps1 - Analysis checkpoint and resume
# ============================================
# Saves intermediate state during long-running analyses
# so operations can resume after failures or interruptions.

$script:CheckpointData = $null
$script:CheckpointPath = $null

function Start-Checkpoint {
    <#
    .SYNOPSIS
    Initializes a new checkpoint session for tracking analysis progress
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$OperationType,
        [string]$Scope = "All"
    )

    $logPath = Get-AppSetting -SettingName "Logging.LogPath"
    if (-not $logPath) { $logPath = "./Logs" }
    if (-not (Test-Path $logPath)) {
        New-Item -Path $logPath -ItemType Directory -Force | Out-Null
    }

    $script:CheckpointPath = Join-Path $logPath "checkpoint_$($OperationType.ToLower()).json"

    $script:CheckpointData = @{
        OperationType   = $OperationType
        Scope           = $Scope
        StartedAt       = (Get-Date).ToString("o")
        LastUpdated     = (Get-Date).ToString("o")
        Phase           = "Initializing"
        CompletedPhases = @()
        ProcessedItems  = @{}
        TotalItems      = @{}
        Status          = "InProgress"
    }

    Save-CheckpointFile
    Write-ActivityLog "Checkpoint started: $OperationType (scope: $Scope)" -Level "Information"
}

function Update-Checkpoint {
    <#
    .SYNOPSIS
    Updates the current checkpoint with progress information
    #>
    param(
        [string]$Phase,
        [string]$ItemKey,
        [int]$ProcessedCount,
        [int]$TotalCount
    )

    if (-not $script:CheckpointData) { return }

    $script:CheckpointData.LastUpdated = (Get-Date).ToString("o")

    if ($Phase) {
        $script:CheckpointData.Phase = $Phase
    }

    if ($ItemKey) {
        $script:CheckpointData.ProcessedItems[$ItemKey] = $ProcessedCount
        if ($TotalCount -gt 0) {
            $script:CheckpointData.TotalItems[$ItemKey] = $TotalCount
        }
    }

    Save-CheckpointFile
}

function Complete-Checkpoint {
    <#
    .SYNOPSIS
    Marks the entire checkpoint as completed and cleans up
    #>
    param(
        [string]$Status = "Completed"
    )

    if (-not $script:CheckpointData) { return }

    $script:CheckpointData.Status = $Status
    $script:CheckpointData.LastUpdated = (Get-Date).ToString("o")
    $script:CheckpointData.CompletedAt = (Get-Date).ToString("o")

    Save-CheckpointFile

    # Clean up checkpoint file on success
    if ($Status -eq "Completed" -and $script:CheckpointPath -and (Test-Path $script:CheckpointPath)) {
        Remove-Item $script:CheckpointPath -Force -ErrorAction SilentlyContinue
    }

    Write-ActivityLog "Checkpoint $Status`: $($script:CheckpointData.OperationType)" -Level "Information"
    $script:CheckpointData = $null
}

function Get-Checkpoint {
    <#
    .SYNOPSIS
    Retrieves a saved checkpoint for potential resume
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$OperationType
    )

    $logPath = Get-AppSetting -SettingName "Logging.LogPath"
    if (-not $logPath) { $logPath = "./Logs" }

    $checkFile = Join-Path $logPath "checkpoint_$($OperationType.ToLower()).json"

    if (Test-Path $checkFile) {
        try {
            $data = Get-Content $checkFile -Raw | ConvertFrom-Json
            if ($data.Status -eq "InProgress") {
                return $data
            }
        }
        catch {
            Write-ActivityLog "Failed to read checkpoint: $($_.Exception.Message)" -Level "Warning"
        }
    }

    return $null
}

function Save-CheckpointFile {
    <#
    .SYNOPSIS
    Internal: persists checkpoint data to disk
    #>
    if (-not $script:CheckpointData -or -not $script:CheckpointPath) { return }

    try {
        $script:CheckpointData | ConvertTo-Json -Depth 5 | Set-Content $script:CheckpointPath -Encoding UTF8
    }
    catch {
        Write-ActivityLog "Failed to save checkpoint: $($_.Exception.Message)" -Level "Warning"
    }
}
