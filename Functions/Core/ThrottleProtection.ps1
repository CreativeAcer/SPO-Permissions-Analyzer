# ============================================
# ThrottleProtection.ps1 - Retry with backoff for PnP calls
# ============================================
# Wraps SharePoint API calls with throttle detection and
# exponential backoff retry logic to handle 429/503 responses.

$script:ThrottleState = @{
    TotalRetries    = 0
    TotalThrottled  = 0
    LastThrottledAt = $null
}

function Invoke-WithThrottleProtection {
    <#
    .SYNOPSIS
    Executes a script block with automatic retry on SharePoint throttling
    .PARAMETER ScriptBlock
    The PnP command(s) to execute
    .PARAMETER OperationName
    Friendly name for logging
    .PARAMETER MaxRetries
    Maximum retry attempts (default 5)
    .PARAMETER InitialBackoffMs
    Starting backoff delay in milliseconds (default 2000)
    #>
    param(
        [Parameter(Mandatory = $true)]
        [scriptblock]$ScriptBlock,
        [string]$OperationName = "SharePoint operation",
        [int]$MaxRetries = 5,
        [int]$InitialBackoffMs = 2000
    )

    $attempt = 0
    $backoffMs = $InitialBackoffMs

    while ($true) {
        try {
            $result = & $ScriptBlock
            return $result
        }
        catch {
            $attempt++
            $errorMsg = $_.Exception.Message

            $isThrottle = $false
            if ($errorMsg -match "429" -or
                $errorMsg -match "503" -or
                $errorMsg -match "throttl" -or
                $errorMsg -match "too many requests" -or
                $errorMsg -match "server is busy" -or
                $errorMsg -match "retry-after") {
                $isThrottle = $true
            }

            if ($isThrottle -and $attempt -le $MaxRetries) {
                $script:ThrottleState.TotalRetries++
                $script:ThrottleState.TotalThrottled++
                $script:ThrottleState.LastThrottledAt = (Get-Date).ToString("o")

                # Check for Retry-After header hint in error message
                $retryAfterMs = $backoffMs
                if ($errorMsg -match "retry-after[:\s]+(\d+)") {
                    $retryAfterMs = [Math]::Max($backoffMs, [int]$Matches[1] * 1000)
                }

                Write-ActivityLog "Throttled on '$OperationName' (attempt $attempt/$MaxRetries). Waiting $($retryAfterMs/1000)s..." -Level "Warning"
                Start-Sleep -Milliseconds $retryAfterMs

                # Exponential backoff with jitter
                $jitter = Get-Random -Minimum 0 -Maximum ([Math]::Max(1, $backoffMs / 4))
                $backoffMs = [Math]::Min($backoffMs * 2 + $jitter, 60000)
            }
            elseif ($attempt -le $MaxRetries -and $errorMsg -match "timeout|timed out") {
                # Retry on timeouts too
                $script:ThrottleState.TotalRetries++
                Write-ActivityLog "Timeout on '$OperationName' (attempt $attempt/$MaxRetries). Retrying in $($backoffMs/1000)s..." -Level "Warning"
                Start-Sleep -Milliseconds $backoffMs
                $backoffMs = [Math]::Min($backoffMs * 2, 60000)
            }
            else {
                # Not a throttle error or max retries exceeded
                if ($attempt -gt $MaxRetries) {
                    Write-ActivityLog "Max retries ($MaxRetries) exceeded for '$OperationName'" -Level "Error"
                }
                throw
            }
        }
    }
}

function Get-ThrottleStats {
    <#
    .SYNOPSIS
    Returns throttling statistics for the current session
    #>
    return $script:ThrottleState
}

function Reset-ThrottleStats {
    <#
    .SYNOPSIS
    Resets throttle counters
    #>
    $script:ThrottleState = @{
        TotalRetries    = 0
        TotalThrottled  = 0
        LastThrottledAt = $null
    }
}
