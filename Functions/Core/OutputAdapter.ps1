# ============================================
# OutputAdapter.ps1 - Console Output Abstraction
# ============================================
# Provides Write-ConsoleOutput and Update-UIAndWait functions
# that can be overridden by the web server or background runspace.
# In web mode, Write-ConsoleOutput writes to the shared operation log.
# In background runspaces, it writes to the synchronized SharedState.

function Write-ConsoleOutput {
    <#
    .SYNOPSIS
    Writes a message to the operation log (web mode default)
    .DESCRIPTION
    This is the default implementation for web mode.
    It forwards messages to Add-OperationLog when the server is running.
    In background runspaces, this function is overridden to write
    directly to the shared operation log.
    #>
    param(
        [string]$Message,
        [switch]$Append,
        [switch]$NewLine = $true,
        [switch]$ForceUpdate
    )
    if ($script:ServerState -and $script:ServerState.OperationLog) {
        [void]$script:ServerState.OperationLog.Add($Message)
    }
}

function Update-UIAndWait {
    <#
    .SYNOPSIS
    No-op in web mode (WPF dispatcher not needed)
    #>
    param([int]$WaitMs = 0)
}
