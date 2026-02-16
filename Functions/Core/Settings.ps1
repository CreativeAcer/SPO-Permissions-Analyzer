# Global settings
$script:AppSettings = @{
    DemoMode = $false
    SharePoint = @{
        TenantUrl = ""
        SiteUrl = ""
    }
    Logging = @{
        LogPath = "./Logs"
        LogLevel = "Information"
    }
}

# Connection state (centralized here instead of scattered across entry points)
$script:SPOConnected = $false
$script:DemoMode = $false

function Initialize-Settings {
    Write-Host "Settings initialized." -ForegroundColor Green
}

function Get-AppSetting {
    param([string]$SettingName)
    
    $pathParts = $SettingName.Split('.')
    $current = $script:AppSettings
    
    foreach ($part in $pathParts) {
        if ($current -is [hashtable] -and $current.ContainsKey($part)) {
            $current = $current[$part]
        } else {
            return $null
        }
    }
    
    return $current
}

function Set-AppSetting {
    param([string]$SettingName, $Value)
    
    $pathParts = $SettingName.Split('.')
    
    if ($pathParts.Length -eq 1) {
        $script:AppSettings[$SettingName] = $Value
    } else {
        $current = $script:AppSettings
        for ($i = 0; $i -lt $pathParts.Length - 1; $i++) {
            $part = $pathParts[$i]
            if (-not $current.ContainsKey($part)) {
                $current[$part] = @{}
            }
            $current = $current[$part]
        }
        $current[$pathParts[-1]] = $Value
    }
}
