# ============================================
# RiskScoring.ps1 - Risk scoring and findings engine
# ============================================
# Evaluates collected SharePoint data against security rules
# and produces scored findings with severity levels.

function Get-RiskAssessment {
    <#
    .SYNOPSIS
    Runs all risk rules against the current data and returns scored findings
    #>
    $metrics = Get-SharePointData -DataType "Metrics"
    $users = Get-SharePointData -DataType "Users"
    $groups = Get-SharePointData -DataType "Groups"
    $roleAssignments = Get-SharePointData -DataType "RoleAssignments"
    $inheritanceItems = Get-SharePointData -DataType "InheritanceItems"
    $sharingLinks = Get-SharePointData -DataType "SharingLinks"

    $findings = [System.Collections.ArrayList]::new()

    # ---- External Access Rules ----

    $externalUsers = @($users | Where-Object { $_.Type -eq "External" -or $_.IsExternal })
    if ($externalUsers.Count -gt 0) {
        $externalEditors = @($externalUsers | Where-Object {
            $_.Permission -in @("Full Control", "Edit", "Contribute")
        })
        if ($externalEditors.Count -gt 0) {
            [void]$findings.Add(@{
                RuleId      = "EXT-001"
                Severity    = "High"
                Category    = "External Access"
                Title       = "External users with edit or higher permissions"
                Description = "$($externalEditors.Count) external user(s) have edit or higher access. Review and restrict to read-only where possible."
                Count       = $externalEditors.Count
                Score       = [Math]::Min($externalEditors.Count * 15, 100)
            })
        }

        $externalAdmins = @($externalUsers | Where-Object { $_.IsSiteAdmin })
        if ($externalAdmins.Count -gt 0) {
            [void]$findings.Add(@{
                RuleId      = "EXT-002"
                Severity    = "Critical"
                Category    = "External Access"
                Title       = "External users with site admin rights"
                Description = "$($externalAdmins.Count) external user(s) are site administrators. This is a significant security risk."
                Count       = $externalAdmins.Count
                Score       = 100
            })
        }

        # Domain diversity
        $domains = @{}
        foreach ($u in $externalUsers) {
            if ($u.Email -and $u.Email -match "@(.+)$") {
                $domain = $Matches[1]
                $domains[$domain] = ($domains[$domain] ?? 0) + 1
            }
        }
        if ($domains.Count -gt 5) {
            [void]$findings.Add(@{
                RuleId      = "EXT-003"
                Severity    = "Medium"
                Category    = "External Access"
                Title       = "External access from many domains"
                Description = "External users come from $($domains.Count) different domains. Consider consolidating external access."
                Count       = $domains.Count
                Score       = [Math]::Min($domains.Count * 5, 70)
            })
        }
    }

    # ---- Sharing Link Rules ----

    $anonymousLinks = @($sharingLinks | Where-Object { $_.LinkType -eq "Anonymous" })
    if ($anonymousLinks.Count -gt 0) {
        $anonymousEditLinks = @($anonymousLinks | Where-Object { $_.AccessLevel -eq "Edit" })
        if ($anonymousEditLinks.Count -gt 0) {
            [void]$findings.Add(@{
                RuleId      = "SHARE-001"
                Severity    = "Critical"
                Category    = "Sharing Links"
                Title       = "Anonymous edit links detected"
                Description = "$($anonymousEditLinks.Count) anonymous link(s) grant edit access. Anyone with the link can modify content without authentication."
                Count       = $anonymousEditLinks.Count
                Score       = 100
            })
        }

        [void]$findings.Add(@{
            RuleId      = "SHARE-002"
            Severity    = "High"
            Category    = "Sharing Links"
            Title       = "Anonymous sharing links exist"
            Description = "$($anonymousLinks.Count) anonymous link(s) allow access without authentication. Review and remove unnecessary links."
            Count       = $anonymousLinks.Count
            Score       = [Math]::Min($anonymousLinks.Count * 20, 90)
        })
    }

    $orgLinks = @($sharingLinks | Where-Object { $_.LinkType -eq "Company-wide" -or $_.LinkType -eq "Organization" })
    if ($orgLinks.Count -gt 10) {
        [void]$findings.Add(@{
            RuleId      = "SHARE-003"
            Severity    = "Medium"
            Category    = "Sharing Links"
            Title       = "Excessive company-wide sharing links"
            Description = "$($orgLinks.Count) company-wide sharing links found. Use specific-people links for sensitive content."
            Count       = $orgLinks.Count
            Score       = [Math]::Min($orgLinks.Count * 3, 60)
        })
    }

    # ---- Permission Rules ----

    $fullControlAssignments = @($roleAssignments | Where-Object { $_.Role -eq "Full Control" })
    if ($fullControlAssignments.Count -gt 5) {
        [void]$findings.Add(@{
            RuleId      = "PERM-001"
            Severity    = "High"
            Category    = "Permissions"
            Title       = "Excessive Full Control assignments"
            Description = "$($fullControlAssignments.Count) Full Control assignments found. Apply least-privilege principle."
            Count       = $fullControlAssignments.Count
            Score       = [Math]::Min($fullControlAssignments.Count * 8, 80)
        })
    }

    $directUserAssignments = @($roleAssignments | Where-Object { $_.PrincipalType -eq "User" })
    if ($directUserAssignments.Count -gt 10) {
        [void]$findings.Add(@{
            RuleId      = "PERM-002"
            Severity    = "Medium"
            Category    = "Permissions"
            Title       = "Many direct user permission assignments"
            Description = "$($directUserAssignments.Count) permissions granted directly to users. Use groups for easier management."
            Count       = $directUserAssignments.Count
            Score       = [Math]::Min($directUserAssignments.Count * 3, 50)
        })
    }

    # ---- Inheritance Rules ----

    $totalItems = @($inheritanceItems).Count
    $brokenItems = @($inheritanceItems | Where-Object {
        $_.HasUniquePermissions -eq $true -or $_.HasUniquePermissions -eq "True"
    })
    if ($totalItems -gt 0) {
        $breakPercentage = [Math]::Round(($brokenItems.Count / $totalItems) * 100, 0)

        if ($breakPercentage -gt 50) {
            [void]$findings.Add(@{
                RuleId      = "INH-001"
                Severity    = "High"
                Category    = "Inheritance"
                Title       = "Majority of items have broken inheritance"
                Description = "$breakPercentage% of scanned items ($($brokenItems.Count)/$totalItems) have unique permissions. Consider consolidating at site level."
                Count       = $brokenItems.Count
                Score       = [Math]::Min($breakPercentage, 85)
            })
        }
        elseif ($breakPercentage -gt 25) {
            [void]$findings.Add(@{
                RuleId      = "INH-002"
                Severity    = "Medium"
                Category    = "Inheritance"
                Title       = "Significant inheritance breaks"
                Description = "$breakPercentage% of scanned items have broken inheritance. Review and consolidate where possible."
                Count       = $brokenItems.Count
                Score       = [Math]::Min($breakPercentage, 60)
            })
        }
    }

    # ---- Group Rules ----

    $emptyGroups = @($groups | Where-Object { ([int]$_.MemberCount) -eq 0 })
    if ($emptyGroups.Count -gt 0) {
        [void]$findings.Add(@{
            RuleId      = "GRP-001"
            Severity    = "Low"
            Category    = "Groups"
            Title       = "Empty groups detected"
            Description = "$($emptyGroups.Count) group(s) have no members. Consider removing unused groups."
            Count       = $emptyGroups.Count
            Score       = [Math]::Min($emptyGroups.Count * 5, 30)
        })
    }

    # Sort findings by score descending
    $sortedFindings = @($findings | Sort-Object { $_.Score } -Descending)

    # Calculate overall risk score (weighted average of top findings, max 100)
    $overallScore = 0
    if ($sortedFindings.Count -gt 0) {
        $topScores = @($sortedFindings | Select-Object -First 5 | ForEach-Object { $_.Score })
        $overallScore = [Math]::Min(($topScores | Measure-Object -Average).Average, 100)
        $overallScore = [Math]::Round($overallScore, 0)
    }

    $riskLevel = switch ($overallScore) {
        { $_ -ge 80 } { "Critical"; break }
        { $_ -ge 60 } { "High"; break }
        { $_ -ge 30 } { "Medium"; break }
        { $_ -gt 0 }  { "Low"; break }
        default        { "None" }
    }

    return @{
        OverallScore = $overallScore
        RiskLevel    = $riskLevel
        TotalFindings = $sortedFindings.Count
        CriticalCount = @($sortedFindings | Where-Object { $_.Severity -eq "Critical" }).Count
        HighCount     = @($sortedFindings | Where-Object { $_.Severity -eq "High" }).Count
        MediumCount   = @($sortedFindings | Where-Object { $_.Severity -eq "Medium" }).Count
        LowCount      = @($sortedFindings | Where-Object { $_.Severity -eq "Low" }).Count
        Findings      = $sortedFindings
    }
}
