function Write-ErrorLog {
    param([string]$Message, [string]$Location = "Unknown")
    
    try {
        $logPath = Get-AppSetting -SettingName "Logging.LogPath"
        if (-not $logPath) { $logPath = "./Logs" }
        
        if (-not (Test-Path $logPath)) {
            New-Item -Path $logPath -ItemType Directory -Force | Out-Null
        }
        
        $logFile = Join-Path $logPath "error_log.txt"
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        $logEntry = "[$timestamp] [$Location] $Message"
        
        Add-Content -Path $logFile -Value $logEntry -Encoding UTF8
        Write-Host $logEntry -ForegroundColor Red
    }
    catch {
        Write-Warning "Failed to write error log: $_"
    }
}

function Write-ActivityLog {
    param([string]$Message, [string]$Level = "Information")

    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        "Information" { Write-Host $logEntry -ForegroundColor White }
        "Warning" { Write-Host $logEntry -ForegroundColor Yellow }
        "Error" { Write-Host $logEntry -ForegroundColor Red }
        default { Write-Host $logEntry }
    }

    # Persist to activity log file
    try {
        $logPath = Get-AppSetting -SettingName "Logging.LogPath"
        if (-not $logPath) { $logPath = "./Logs" }
        if (-not (Test-Path $logPath)) {
            New-Item -Path $logPath -ItemType Directory -Force | Out-Null
        }
        $logFile = Join-Path $logPath "activity_log.txt"
        Add-Content -Path $logFile -Value $logEntry -Encoding UTF8
    }
    catch { }
}
