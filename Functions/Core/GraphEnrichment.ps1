# ============================================
# GraphEnrichment.ps1 - External user enrichment via Microsoft Graph
# ============================================
# Enriches external user data with additional properties from
# Microsoft Graph: UserType, AccountEnabled, LastSignIn, CreatedDateTime.
# Uses PnP PowerShell's Invoke-PnPGraphMethod (no extra modules needed).

function Invoke-ExternalUserEnrichment {
    <#
    .SYNOPSIS
    Enriches external users in the data store with Graph API data
    .DESCRIPTION
    For each external user, queries Microsoft Graph to add:
    - UserType (Guest/Member)
    - AccountEnabled (true/false)
    - LastSignIn timestamp
    - CreatedDateTime
    - DisplayName from Graph
    Requires User.Read.All delegated permission.
    #>
    param(
        [switch]$Force
    )

    $users = Get-SharePointData -DataType "Users"
    if (-not $users -or $users.Count -eq 0) {
        Write-ActivityLog "No users to enrich" -Level "Warning"
        return
    }

    $externalUsers = @($users | Where-Object {
        $_.Type -eq "External" -or $_.IsExternal -eq $true
    })

    if ($externalUsers.Count -eq 0) {
        Write-ActivityLog "No external users found for enrichment" -Level "Information"
        return
    }

    Write-ActivityLog "Starting Graph enrichment for $($externalUsers.Count) external users" -Level "Information"

    # Test Graph access first
    if (-not (Test-GraphAccess)) {
        Write-ActivityLog "Graph API not available - skipping enrichment" -Level "Warning"
        return
    }

    $enrichedCount = 0
    $failedCount = 0

    foreach ($user in $externalUsers) {
        # Skip already enriched users unless forced
        if (-not $Force -and $user.GraphEnriched) { continue }

        try {
            $graphData = Get-GraphUserData -Email $user.Email -LoginName $user.LoginName
            if ($graphData) {
                $user.GraphUserType = $graphData.UserType
                $user.GraphAccountEnabled = $graphData.AccountEnabled
                $user.GraphLastSignIn = $graphData.LastSignIn
                $user.GraphCreatedDate = $graphData.CreatedDateTime
                $user.GraphDisplayName = $graphData.DisplayName
                $user.GraphEnriched = $true
                $enrichedCount++
            }
            else {
                $user.GraphEnriched = $false
                $failedCount++
            }
        }
        catch {
            Write-ActivityLog "Failed to enrich user $($user.Email): $($_.Exception.Message)" -Level "Warning"
            $user.GraphEnriched = $false
            $failedCount++
        }

        # Small delay to avoid throttling
        Start-Sleep -Milliseconds 100
    }

    Write-ActivityLog "Graph enrichment complete: $enrichedCount enriched, $failedCount failed" -Level "Information"

    return @{
        TotalExternal = $externalUsers.Count
        Enriched      = $enrichedCount
        Failed        = $failedCount
    }
}

function Test-GraphAccess {
    <#
    .SYNOPSIS
    Tests if Graph API is accessible via the current PnP connection
    #>
    try {
        $result = Invoke-PnPGraphMethod -Url "v1.0/me" -Method Get -ErrorAction Stop
        return $null -ne $result
    }
    catch {
        Write-ActivityLog "Graph access test failed: $($_.Exception.Message)" -Level "Warning"
        return $false
    }
}

function Get-GraphUserData {
    <#
    .SYNOPSIS
    Retrieves user data from Microsoft Graph by email or UPN
    #>
    param(
        [string]$Email,
        [string]$LoginName
    )

    # Extract UPN from login name if available
    $upn = $null
    if ($LoginName -match "membership\|(.+)$") {
        $upn = $Matches[1]
    }

    $searchEmail = if ($upn) { $upn } elseif ($Email -and $Email -ne "N/A") { $Email } else { return $null }

    # Try to find user by mail or UPN
    $filter = "mail eq '$searchEmail' or userPrincipalName eq '$searchEmail'"
    $select = "id,displayName,userType,accountEnabled,createdDateTime,signInActivity"

    try {
        $result = Invoke-PnPGraphMethod -Url "v1.0/users?`$filter=$filter&`$select=$select" -Method Get -ErrorAction Stop

        if ($result -and $result.value -and $result.value.Count -gt 0) {
            $graphUser = $result.value[0]

            $lastSignIn = $null
            if ($graphUser.signInActivity -and $graphUser.signInActivity.lastSignInDateTime) {
                $lastSignIn = $graphUser.signInActivity.lastSignInDateTime
            }

            return @{
                DisplayName     = $graphUser.displayName
                UserType        = $graphUser.userType
                AccountEnabled  = $graphUser.accountEnabled
                CreatedDateTime = $graphUser.createdDateTime
                LastSignIn      = $lastSignIn
            }
        }
    }
    catch {
        # Try direct lookup by UPN/email
        try {
            $encodedEmail = [System.Uri]::EscapeDataString($searchEmail)
            $result = Invoke-PnPGraphMethod -Url "v1.0/users/$encodedEmail`?`$select=$select" -Method Get -ErrorAction Stop

            if ($result) {
                $lastSignIn = $null
                if ($result.signInActivity -and $result.signInActivity.lastSignInDateTime) {
                    $lastSignIn = $result.signInActivity.lastSignInDateTime
                }

                return @{
                    DisplayName     = $result.displayName
                    UserType        = $result.userType
                    AccountEnabled  = $result.accountEnabled
                    CreatedDateTime = $result.createdDateTime
                    LastSignIn      = $lastSignIn
                }
            }
        }
        catch {
            return $null
        }
    }

    return $null
}

function Get-EnrichmentSummary {
    <#
    .SYNOPSIS
    Returns a summary of enrichment data for external users
    #>
    $users = Get-SharePointData -DataType "Users"
    $external = @($users | Where-Object { $_.Type -eq "External" -or $_.IsExternal })
    $enriched = @($external | Where-Object { $_.GraphEnriched -eq $true })

    $disabledAccounts = @($enriched | Where-Object { $_.GraphAccountEnabled -eq $false })
    $guestUsers = @($enriched | Where-Object { $_.GraphUserType -eq "Guest" })

    # Stale accounts (no sign-in in 90 days)
    $staleAccounts = @()
    $cutoff = (Get-Date).AddDays(-90)
    foreach ($u in $enriched) {
        if ($u.GraphLastSignIn) {
            try {
                $signInDate = [DateTime]::Parse($u.GraphLastSignIn)
                if ($signInDate -lt $cutoff) {
                    $staleAccounts += $u
                }
            }
            catch { }
        }
        elseif ($u.GraphCreatedDate) {
            # No sign-in recorded, check if created before cutoff
            try {
                $createdDate = [DateTime]::Parse($u.GraphCreatedDate)
                if ($createdDate -lt $cutoff) {
                    $staleAccounts += $u
                }
            }
            catch { }
        }
    }

    return @{
        TotalExternal    = $external.Count
        EnrichedCount    = $enriched.Count
        DisabledAccounts = $disabledAccounts.Count
        GuestUsers       = $guestUsers.Count
        StaleAccounts    = $staleAccounts.Count
        StaleAccountList = $staleAccounts
    }
}
