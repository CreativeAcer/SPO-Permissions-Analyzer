# ============================================
# SiteCollector.ps1 - SharePoint Site Enumeration
# ============================================
# Retrieves real SharePoint sites with storage data.
# Runs inside background runspaces (Start-BackgroundOperation)
# where Write-ConsoleOutput writes to the shared operation log.

function Get-RealSites-DataDriven {
    <#
    .SYNOPSIS
    Retrieves real SharePoint sites with proper storage data
    #>
    try {
        # Clear previous sites data and set context
        Clear-SharePointData -DataType "Sites"
        Set-SharePointOperationContext -OperationType "Sites Analysis"
        Start-Checkpoint -OperationType "SitesAnalysis" -Scope "Tenant"
        Reset-ThrottleStats

        # Start audit session
        Start-AuditSession -OperationType "SitesAnalysis" -ScanScope "Tenant"

        Write-ConsoleOutput "SHAREPOINT SITES ANALYSIS"
        Write-ConsoleOutput "====================================================="
        Write-ConsoleOutput "Started at: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')"
        Write-ConsoleOutput ""

        Write-ConsoleOutput "Using modern PnP PowerShell for site enumeration..."
        Write-ConsoleOutput "Attempting tenant-level site enumeration..."

        $sites = @()
        $requiresAdmin = $false

        # Get tenant URL and client ID for reconnection
        $tenantUrl = Get-AppSetting -SettingName "SharePoint.TenantUrl"
        $adminUrl = $tenantUrl -replace "\.sharepoint\.com", "-admin.sharepoint.com"
        $clientId = Get-AppSetting -SettingName "SharePoint.ClientId"

        # Try to get sites with full details using the existing connection
        try {
            Write-ConsoleOutput "Scanning tenant for site collections..."

            # First try Get-PnPTenantSite on the existing connection (may need admin)
            try {
                Write-ConsoleOutput "Attempting to enumerate tenant sites..."
                $sites = Invoke-WithThrottleProtection -OperationName "Get-PnPTenantSite" -ScriptBlock {
                    Get-PnPTenantSite -ErrorAction Stop
                }
                Write-ActivityLog "Retrieved $($sites.Count) sites from tenant"
                Write-ConsoleOutput "Successfully retrieved $($sites.Count) sites"
            }
            catch {
                Write-ActivityLog "Standard enumeration failed, trying admin connection: $($_.Exception.Message)"
                Write-ConsoleOutput "Standard access limited, trying admin center..."

                # Fallback to admin connection
                $token = $null
                try { $token = Get-PnPAccessToken -ErrorAction SilentlyContinue } catch { }

                if ($token) {
                    Connect-PnPOnline -Url $adminUrl -AccessToken $token -ErrorAction Stop
                } else {
                    Connect-PnPOnline -Url $adminUrl -ClientId $clientId -Interactive -ErrorAction Stop
                }
                $sites = Invoke-WithThrottleProtection -OperationName "Get-PnPTenantSite (admin)" -ScriptBlock {
                    Get-PnPTenantSite -Detailed -ErrorAction Stop
                }
                $requiresAdmin = $true
                Write-ActivityLog "Retrieved $($sites.Count) sites from admin center"
                Write-ConsoleOutput "Successfully retrieved $($sites.Count) sites with storage details"
            }
        }
        catch {
            Write-ActivityLog "Tenant-level enumeration failed: $($_.Exception.Message)"
            Write-ConsoleOutput "Tenant-level access limited, using fallback method..."

            # Final fallback - get current site only
            try {
                $currentWeb = Get-PnPWeb -ErrorAction Stop
                $currentSite = Get-PnPSite -Includes Usage -ErrorAction Stop

                $siteObj = [PSCustomObject]@{
                    Title = $currentWeb.Title
                    Url = $currentWeb.Url
                    StorageUsageCurrent = if ($currentSite.Usage -and $currentSite.Usage.Storage) {
                        [math]::Round($currentSite.Usage.Storage / 1MB, 0)
                    } else { 0 }
                    StorageQuota = if ($currentSite.Usage -and $currentSite.Usage.StoragePercentageUsed) {
                        [math]::Round($currentSite.Usage.StoragePercentageUsed * 100, 0)
                    } else { 0 }
                    Template = $currentWeb.WebTemplate
                    LastContentModifiedDate = $currentWeb.LastItemModifiedDate
                    Owner = "Current User"
                }

                $sites = @($siteObj)
                Write-ConsoleOutput "Found current site only"
            }
            catch {
                throw "Unable to retrieve any sites. SharePoint Administrator permissions may be required."
            }
        }

        Write-ConsoleOutput ""
        Write-ConsoleOutput "SITES DISCOVERY RESULTS"
        Write-ConsoleOutput "Sites Found: $($sites.Count)"
        Write-ConsoleOutput ""

        # Process and store each site with proper storage data
        $siteCounter = 0
        foreach ($site in $sites | Select-Object -First 25) {
            $siteCounter++

            # Extract storage value properly based on the object type
            $storageValue = 0

            if ($null -ne $site.StorageUsageCurrent) {
                $storageValue = [int]$site.StorageUsageCurrent
            }
            elseif ($site.Usage -and $site.Usage.Storage) {
                $storageValue = [math]::Round($site.Usage.Storage / 1MB, 0)
            }
            elseif ($site.Url) {
                try {
                    $token = $null
                    try { $token = Get-PnPAccessToken -ErrorAction SilentlyContinue } catch { }
                    if ($token) {
                        Connect-PnPOnline -Url $site.Url -AccessToken $token -ErrorAction SilentlyContinue
                    }
                    $siteDetail = Get-PnPSite -Includes Usage -ErrorAction SilentlyContinue
                    if ($siteDetail -and $siteDetail.Usage -and $siteDetail.Usage.Storage) {
                        $storageValue = [math]::Round($siteDetail.Usage.Storage / 1MB, 0)
                    }
                }
                catch {
                    Write-ActivityLog "Could not get storage for site: $($site.Url)" -Level "Warning"
                }
            }

            $siteData = @{
                Title = if ($site.Title) { $site.Title } else { "Site $siteCounter" }
                Url = if ($site.Url) { $site.Url } else { "N/A" }
                Owner = if ($site.Owner) { $site.Owner } elseif ($site.SiteOwnerEmail) { $site.SiteOwnerEmail } else { "N/A" }
                Storage = $storageValue.ToString()
                StorageQuota = if ($site.StorageQuota) { $site.StorageQuota } else { 0 }
                Template = if ($site.Template) { $site.Template } else { "N/A" }
                LastModified = if ($site.LastContentModifiedDate) {
                    $site.LastContentModifiedDate.ToString("yyyy-MM-dd")
                } else { "N/A" }
                IsHubSite = if ($site.IsHubSite) { $site.IsHubSite } else { $false }
                UserCount = 0
                GroupCount = 0
            }

            Add-SharePointSite -SiteData $siteData

            Write-ConsoleOutput "SITE #${siteCounter}: $($siteData.Title)"
            Write-ConsoleOutput "   URL: $($siteData.Url)"
            Write-ConsoleOutput "   Owner: $($siteData.Owner)"
            Write-ConsoleOutput "   Storage: $storageValue MB"
            if ($siteData.StorageQuota -gt 0) {
                $quotaMB = [math]::Round($siteData.StorageQuota / 1MB, 0)
                $percentUsed = if ($quotaMB -gt 0) {
                    [math]::Round(($storageValue / $quotaMB) * 100, 1)
                } else { 0 }
                Write-ConsoleOutput "   Storage Quota: $quotaMB MB ($percentUsed% used)"
            }
            if ($siteData.Template -ne "N/A") { Write-ConsoleOutput "   Template: $($siteData.Template)" }
            if ($siteData.LastModified -ne "N/A") { Write-ConsoleOutput "   Last Modified: $($siteData.LastModified)" }
        }

        if ($sites.Count -gt 25) {
            Write-ConsoleOutput "... and $($sites.Count - 25) more sites"
            Write-ConsoleOutput ""
        }

        # Show storage summary
        $totalStorage = 0
        $sitesWithStorage = 0
        $allSites = Get-SharePointData -DataType "Sites"
        foreach ($s in $allSites) {
            $storage = [int]$s["Storage"]
            if ($storage -gt 0) {
                $totalStorage += $storage
                $sitesWithStorage++
            }
        }

        if ($sitesWithStorage -gt 0) {
            $totalGB = [math]::Round($totalStorage / 1024, 2)
            $avgMB = [math]::Round($totalStorage / $sitesWithStorage, 0)
            Write-ConsoleOutput ""
            Write-ConsoleOutput "STORAGE SUMMARY:"
            Write-ConsoleOutput "   Total Storage Used: $totalStorage MB ($totalGB GB)"
            Write-ConsoleOutput "   Average per Site: $avgMB MB"
            Write-ConsoleOutput "   Sites with Storage Data: $sitesWithStorage/$($sites.Count)"
        }

        if (-not $requiresAdmin -and $sitesWithStorage -eq 0) {
            Write-ConsoleOutput ""
            Write-ConsoleOutput "Note: Storage data requires SharePoint Administrator permissions"
        }

        Write-ConsoleOutput ""
        Write-ConsoleOutput "Site enumeration completed successfully!"

        Write-ActivityLog "Sites operation completed with $($sites.Count) sites" -Level "Information"
        Complete-Checkpoint -Status "Completed"
        Write-AuditEvent -EventType "DataCollection" -Detail "Sites analysis complete: $($sites.Count) sites"
        Complete-AuditSession -Status "Completed"

        # Log throttle stats if any retries occurred
        $throttleStats = Get-ThrottleStats
        if ($throttleStats.TotalRetries -gt 0) {
            Write-ConsoleOutput "Throttle protection: $($throttleStats.TotalRetries) retries, $($throttleStats.TotalThrottled) throttle events"
        }
    }
    catch {
        Write-ErrorLog -Message $_.Exception.Message -Location "Get-RealSites-DataDriven"
        Complete-Checkpoint -Status "Failed"
        Write-AuditEvent -EventType "Error" -Detail $_.Exception.Message
        Complete-AuditSession -Status "Failed"
        Write-ConsoleOutput ""
        Write-ConsoleOutput "ERROR: Site enumeration failed"
        Write-ConsoleOutput "Error Details: $($_.Exception.Message)"
    }
}
