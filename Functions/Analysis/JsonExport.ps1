# ============================================
# JsonExport.ps1 - Structured JSON output with stable schema
# ============================================
# Exports SharePoint analysis data as JSON with a versioned,
# stable schema suitable for governance pipelines and automation.

$script:JsonSchemaVersion = "1.0.0"

function Export-GovernanceJson {
    <#
    .SYNOPSIS
    Exports all analysis data as a structured JSON file with a stable schema
    .PARAMETER OutputPath
    Directory to write the JSON file. Defaults to ./Reports/Generated
    .PARAMETER IncludeMetadata
    Include scan metadata (timestamps, tool version, tenant info)
    #>
    param(
        [string]$OutputPath,
        [switch]$IncludeMetadata = $true
    )

    if (-not $OutputPath) {
        $OutputPath = "./Reports/Generated"
    }
    if (-not (Test-Path $OutputPath)) {
        New-Item -Path $OutputPath -ItemType Directory -Force | Out-Null
    }

    $timestamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $filePath = Join-Path $OutputPath "spo_governance_$timestamp.json"

    $report = Build-GovernanceReport -IncludeMetadata:$IncludeMetadata

    $report | ConvertTo-Json -Depth 15 | Set-Content $filePath -Encoding UTF8

    Write-ActivityLog "Governance JSON exported: $filePath" -Level "Information"
    return $filePath
}

function Build-GovernanceReport {
    <#
    .SYNOPSIS
    Builds the structured governance report object
    #>
    param(
        [switch]$IncludeMetadata = $true
    )

    $metrics = Get-SharePointData -DataType "Metrics"
    $sites = Get-SharePointData -DataType "Sites"
    $users = Get-SharePointData -DataType "Users"
    $groups = Get-SharePointData -DataType "Groups"
    $roleAssignments = Get-SharePointData -DataType "RoleAssignments"
    $inheritanceItems = Get-SharePointData -DataType "InheritanceItems"
    $sharingLinks = Get-SharePointData -DataType "SharingLinks"

    $report = [ordered]@{
        schemaVersion = $script:JsonSchemaVersion
        exportedAt    = (Get-Date).ToString("o")
    }

    # Metadata section
    if ($IncludeMetadata) {
        $report.metadata = [ordered]@{
            toolName          = "SPO-Permissions-Analyzer"
            toolVersion       = "1.1.0"
            tenantUrl         = (Get-AppSetting -SettingName "SharePoint.TenantUrl")
            scanTimestamp     = (Get-Date).ToString("o")
            powerShellVersion = $PSVersionTable.PSVersion.ToString()
            hostName          = $env:COMPUTERNAME
            demoMode          = [bool]$script:DemoMode
        }
    }

    # Summary metrics
    $report.summary = [ordered]@{
        totalSites           = [int]($metrics.TotalSites)
        totalUsers           = [int]($metrics.TotalUsers)
        totalGroups          = [int]($metrics.TotalGroups)
        externalUsers        = [int]($metrics.ExternalUsers)
        totalRoleAssignments = [int]($metrics.TotalRoleAssignments)
        inheritanceBreaks    = [int]($metrics.InheritanceBreaks)
        totalSharingLinks    = [int]($metrics.TotalSharingLinks)
    }

    # Sites
    $report.sites = @(foreach ($s in $sites) {
        [ordered]@{
            title        = $s.Title
            url          = $s.Url
            owner        = $s.Owner
            storageMB    = [int]($s.Storage)
            template     = if ($s.Template) { $s.Template } else { $null }
            lastModified = if ($s.LastModified) { $s.LastModified } else { $null }
            usageLevel   = if ($s.UsageLevel) { $s.UsageLevel } else { $null }
        }
    })

    # Users
    $report.users = @(foreach ($u in $users) {
        [ordered]@{
            name         = $u.Name
            email        = $u.Email
            type         = $u.Type
            permission   = $u.Permission
            isSiteAdmin  = [bool]($u.IsSiteAdmin)
            isExternal   = [bool]($u.IsExternal -or $u.Type -eq "External")
            loginName    = if ($u.LoginName) { $u.LoginName } else { $null }
        }
    })

    # Groups
    $report.groups = @(foreach ($g in $groups) {
        [ordered]@{
            name        = $g.Name
            memberCount = [int]($g.MemberCount)
            permission  = $g.Permission
            description = if ($g.Description) { $g.Description } else { $null }
        }
    })

    # Role assignments
    $report.roleAssignments = @(foreach ($ra in $roleAssignments) {
        [ordered]@{
            principal     = $ra.Principal
            principalType = $ra.PrincipalType
            role          = $ra.Role
            scope         = $ra.Scope
            scopeUrl      = $ra.ScopeUrl
            siteTitle     = if ($ra.SiteTitle) { $ra.SiteTitle } else { $null }
        }
    })

    # Inheritance items
    $report.inheritance = @(foreach ($i in $inheritanceItems) {
        [ordered]@{
            title                = $i.Title
            type                 = $i.Type
            url                  = $i.Url
            hasUniquePermissions = [bool]($i.HasUniquePermissions)
            parentUrl            = if ($i.ParentUrl) { $i.ParentUrl } else { $null }
            roleAssignmentCount  = [int]($i.RoleAssignmentCount)
            siteTitle            = if ($i.SiteTitle) { $i.SiteTitle } else { $null }
        }
    })

    # Sharing links
    $report.sharingLinks = @(foreach ($l in $sharingLinks) {
        [ordered]@{
            linkType    = $l.LinkType
            accessLevel = $l.AccessLevel
            memberCount = [int]($l.MemberCount)
            siteTitle   = if ($l.SiteTitle) { $l.SiteTitle } else { $null }
            createdDate = if ($l.CreatedDate) { $l.CreatedDate } else { $null }
        }
    })

    return $report
}

