# ============================================
# DemoDataGenerator.ps1 - Demo/Sample Data Generation
# ============================================
# Generates realistic demo data for testing the application
# without a live SharePoint connection.

# Module-level type map (shared across handlers)
$script:DataTypeMap = @{
    "sites"           = "Sites"
    "users"           = "Users"
    "groups"          = "Groups"
    "permissions"     = "Permissions"
    "roleassignments" = "RoleAssignments"
    "inheritance"     = "InheritanceItems"
    "sharinglinks"    = "SharingLinks"
}

function New-DemoData {
    <#
    .SYNOPSIS
    Generates a complete set of demo data and stores it in the data manager
    #>

    Clear-SharePointData -DataType "All"

    # Generate demo sites
    Set-SharePointOperationContext -OperationType "Demo - Sites"
    $demoSites = @(
        @{Title="Team Collaboration Site"; Url="https://contoso.sharepoint.com/sites/teamsite"; Owner="admin@contoso.com"; Storage="750"; Template="STS#3"; LastModified="2025-01-15"; UserCount=25; GroupCount=8},
        @{Title="HR Portal"; Url="https://contoso.sharepoint.com/sites/hr"; Owner="hr.admin@contoso.com"; Storage="1200"; Template="SITEPAGEPUBLISHING#0"; LastModified="2025-02-01"; UserCount=45; GroupCount=12},
        @{Title="Marketing Hub"; Url="https://contoso.sharepoint.com/sites/marketing"; Owner="marketing@contoso.com"; Storage="2500"; Template="STS#3"; LastModified="2025-01-28"; UserCount=30; GroupCount=6},
        @{Title="Finance Department"; Url="https://contoso.sharepoint.com/sites/finance"; Owner="finance.lead@contoso.com"; Storage="450"; Template="STS#3"; LastModified="2025-01-10"; UserCount=15; GroupCount=4},
        @{Title="Executive Dashboard"; Url="https://contoso.sharepoint.com/sites/exec"; Owner="ceo@contoso.com"; Storage="200"; Template="SITEPAGEPUBLISHING#0"; LastModified="2024-12-15"; UserCount=8; GroupCount=3}
    )
    foreach ($site in $demoSites) { Add-SharePointSite -SiteData $site }
    Write-ConsoleOutput "Added $($demoSites.Count) demo sites"

    # Generate demo users (with diverse external users for risk testing)
    Set-SharePointOperationContext -OperationType "Demo - Permissions"
    $demoUsers = @(
        # Internal users
        @{Name="John Doe"; Email="john.doe@contoso.com"; Type="Internal"; Permission="Full Control"; IsSiteAdmin=$true; LoginName="i:0#.f|membership|john.doe@contoso.com"},
        @{Name="Jane Smith"; Email="jane.smith@contoso.com"; Type="Internal"; Permission="Edit"; IsSiteAdmin=$false},
        @{Name="Mike Johnson"; Email="mike.j@contoso.com"; Type="Internal"; Permission="Read"; IsSiteAdmin=$false},
        @{Name="Sarah Wilson"; Email="sarah.w@contoso.com"; Type="Internal"; Permission="Contribute"; IsSiteAdmin=$false},
        @{Name="David Brown"; Email="david.b@contoso.com"; Type="Internal"; Permission="Full Control"; IsSiteAdmin=$true},
        @{Name="Emily Chen"; Email="emily.c@contoso.com"; Type="Internal"; Permission="Edit"; IsSiteAdmin=$false},
        @{Name="Alex Kumar"; Email="alex.k@contoso.com"; Type="Internal"; Permission="Read"; IsSiteAdmin=$false},
        @{Name="Lisa Anderson"; Email="lisa.a@contoso.com"; Type="Internal"; Permission="Contribute"; IsSiteAdmin=$false},
        @{Name="Robert Taylor"; Email="robert.t@contoso.com"; Type="Internal"; Permission="Full Control"; IsSiteAdmin=$false},
        @{Name="Michelle Lee"; Email="michelle.l@contoso.com"; Type="Internal"; Permission="Edit"; IsSiteAdmin=$false},

        # External users - triggers EXT-001 (High - External with Edit+)
        @{Name="External Partner"; Email="partner@external.com"; Type="External"; Permission="Read"; IsExternal=$true},
        @{Name="Guest Reviewer"; Email="reviewer@partner.org"; Type="External"; Permission="Read"; IsExternal=$true},
        @{Name="Contractor A"; Email="contractor.a@vendor.com"; Type="External"; Permission="Edit"; IsExternal=$true},
        @{Name="Consultant Smith"; Email="j.smith@consulting-firm.com"; Type="External"; Permission="Full Control"; IsExternal=$true},
        @{Name="External Dev"; Email="developer@techpartner.io"; Type="External"; Permission="Contribute"; IsExternal=$true},
        @{Name="Agency Designer"; Email="design@agency.co"; Type="External"; Permission="Edit"; IsExternal=$true},
        @{Name="Vendor Contact"; Email="sales@vendor-corp.net"; Type="External"; Permission="Edit"; IsExternal=$true},

        # External admin - triggers EXT-002 (Critical - External Site Admin)
        @{Name="External IT Partner"; Email="admin@it-partner.com"; Type="External"; Permission="Full Control"; IsSiteAdmin=$true; IsExternal=$true},

        # More external users from diverse domains - triggers EXT-003 (Medium - Many domains)
        @{Name="Auditor Jones"; Email="jones@audit-firm.biz"; Type="External"; Permission="Read"; IsExternal=$true},
        @{Name="Legal Advisor"; Email="legal@lawfirm.legal"; Type="External"; Permission="Read"; IsExternal=$true},
        @{Name="Marketing Guest"; Email="guest@marketing-agency.co.uk"; Type="External"; Permission="Contribute"; IsExternal=$true},
        @{Name="Freelancer"; Email="freelance@personal-domain.me"; Type="External"; Permission="Edit"; IsExternal=$true}
    )
    foreach ($user in $demoUsers) { Add-SharePointUser -UserData $user }
    Write-ConsoleOutput "Added $($demoUsers.Count) demo users (including $(@($demoUsers | Where-Object {$_.IsExternal}).Count) external)"

    # Generate demo groups (includes empty groups for GRP-001)
    $demoGroups = @(
        @{Name="Site Owners"; MemberCount=3; Permission="Full Control"; Description="Owners of the site"},
        @{Name="Site Members"; MemberCount=12; Permission="Edit"; Description="Members with edit access"},
        @{Name="Site Visitors"; MemberCount=25; Permission="Read"; Description="Visitors with read access"},
        @{Name="HR Team"; MemberCount=8; Permission="Contribute"; Description="Human Resources team"},
        @{Name="IT Admins"; MemberCount=4; Permission="Full Control"; Description="IT administrators"},
        @{Name="Marketing Team"; MemberCount=15; Permission="Edit"; Description="Marketing department"},
        # Empty groups - triggers GRP-001 (Low - Empty groups)
        @{Name="Legacy Project Team"; MemberCount=0; Permission="Edit"; Description="Old project team - no longer used"},
        @{Name="Temp Contractors Group"; MemberCount=0; Permission="Contribute"; Description="Temporary group created for contractors"},
        @{Name="Archive Access"; MemberCount=0; Permission="Read"; Description="Empty archive group"}
    )
    foreach ($group in $demoGroups) { Add-SharePointGroup -GroupData $group }
    Write-ConsoleOutput "Added $($demoGroups.Count) demo groups (including $(@($demoGroups | Where-Object {$_.MemberCount -eq 0}).Count) empty)"

    # Generate demo role assignments (includes many Full Control and direct user assignments)
    $demoRoles = @(
        # Group assignments
        @{Principal="Site Owners"; PrincipalType="SharePoint Group"; Role="Full Control"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
        @{Principal="Site Members"; PrincipalType="SharePoint Group"; Role="Edit"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
        @{Principal="Site Visitors"; PrincipalType="SharePoint Group"; Role="Read"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
        @{Principal="HR Team"; PrincipalType="Security Group"; Role="Contribute"; Scope="Library"; ScopeUrl="/sites/hr/Shared Documents"; SiteTitle="HR Portal"},
        @{Principal="Finance Group"; PrincipalType="Security Group"; Role="Read"; Scope="List"; ScopeUrl="/sites/finance/Lists/Budget"; SiteTitle="Finance Department"},
        @{Principal="IT Admins"; PrincipalType="Security Group"; Role="Full Control"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
        @{Principal="Marketing Team"; PrincipalType="Security Group"; Role="Edit"; Scope="Library"; ScopeUrl="/sites/marketing/Assets"; SiteTitle="Marketing Hub"},

        # Direct user assignments - triggers PERM-002 (Medium - >10 direct user assignments)
        @{Principal="John Doe"; PrincipalType="User"; Role="Full Control"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
        @{Principal="Jane Smith"; PrincipalType="User"; Role="Edit"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
        @{Principal="Mike Johnson"; PrincipalType="User"; Role="Read"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/teamsite"; SiteTitle="Team Collaboration Site"},
        @{Principal="External Partner"; PrincipalType="User"; Role="Read"; Scope="Library"; ScopeUrl="/sites/teamsite/Shared Documents"; SiteTitle="Team Collaboration Site"},
        @{Principal="Contractor A"; PrincipalType="User"; Role="Edit"; Scope="Library"; ScopeUrl="/sites/teamsite/Project Files"; SiteTitle="Team Collaboration Site"},
        @{Principal="Sarah Wilson"; PrincipalType="User"; Role="Contribute"; Scope="Library"; ScopeUrl="/sites/hr/Policies"; SiteTitle="HR Portal"},
        @{Principal="David Brown"; PrincipalType="User"; Role="Full Control"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/hr"; SiteTitle="HR Portal"},
        @{Principal="Emily Chen"; PrincipalType="User"; Role="Edit"; Scope="Library"; ScopeUrl="/sites/marketing/Assets"; SiteTitle="Marketing Hub"},
        @{Principal="Alex Kumar"; PrincipalType="User"; Role="Read"; Scope="List"; ScopeUrl="/sites/finance/Lists/Budget"; SiteTitle="Finance Department"},
        @{Principal="Lisa Anderson"; PrincipalType="User"; Role="Contribute"; Scope="Library"; ScopeUrl="/sites/teamsite/Shared Documents"; SiteTitle="Team Collaboration Site"},
        @{Principal="Robert Taylor"; PrincipalType="User"; Role="Full Control"; Scope="Library"; ScopeUrl="/sites/finance/Shared Documents"; SiteTitle="Finance Department"},
        @{Principal="Michelle Lee"; PrincipalType="User"; Role="Edit"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/marketing"; SiteTitle="Marketing Hub"},
        @{Principal="External Dev"; PrincipalType="User"; Role="Contribute"; Scope="Library"; ScopeUrl="/sites/teamsite/Project Files"; SiteTitle="Team Collaboration Site"},
        @{Principal="Agency Designer"; PrincipalType="User"; Role="Edit"; Scope="Library"; ScopeUrl="/sites/marketing/Assets"; SiteTitle="Marketing Hub"},

        # Excessive Full Control assignments - triggers PERM-001 (High - >5 Full Control assignments)
        @{Principal="Consultant Smith"; PrincipalType="User"; Role="Full Control"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/exec"; SiteTitle="Executive Dashboard"},
        @{Principal="External IT Partner"; PrincipalType="User"; Role="Full Control"; Scope="Site"; ScopeUrl="https://contoso.sharepoint.com/sites/finance"; SiteTitle="Finance Department"},
        @{Principal="Vendor Contact"; PrincipalType="User"; Role="Full Control"; Scope="Library"; ScopeUrl="/sites/teamsite/Shared Documents"; SiteTitle="Team Collaboration Site"},
        @{Principal="Freelancer"; PrincipalType="User"; Role="Full Control"; Scope="Library"; ScopeUrl="/sites/marketing/Assets"; SiteTitle="Marketing Hub"}
    )
    foreach ($ra in $demoRoles) { Add-SharePointRoleAssignment -RoleData $ra }
    $fullControlCount = @($demoRoles | Where-Object {$_.Role -eq "Full Control"}).Count
    $directUserCount = @($demoRoles | Where-Object {$_.PrincipalType -eq "User"}).Count
    Write-ConsoleOutput "Added $($demoRoles.Count) role assignments ($fullControlCount Full Control, $directUserCount direct user assignments)"

    # Generate demo inheritance items
    $demoInheritance = @(
        @{Title="Team Collaboration Site"; Type="Site"; Url="https://contoso.sharepoint.com/sites/teamsite"; HasUniquePermissions=$true; ParentUrl="N/A"; RoleAssignmentCount=7; SiteTitle="Team Collaboration Site"},
        @{Title="Shared Documents"; Type="Document Library"; Url="/sites/teamsite/Shared Documents"; HasUniquePermissions=$true; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=4; SiteTitle="Team Collaboration Site"},
        @{Title="Site Pages"; Type="Document Library"; Url="/sites/teamsite/SitePages"; HasUniquePermissions=$false; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"},
        @{Title="Project Files"; Type="Document Library"; Url="/sites/teamsite/Project Files"; HasUniquePermissions=$true; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=5; SiteTitle="Team Collaboration Site"},
        @{Title="Marketing Assets"; Type="Document Library"; Url="/sites/marketing/Assets"; HasUniquePermissions=$true; ParentUrl="https://contoso.sharepoint.com/sites/marketing"; RoleAssignmentCount=3; SiteTitle="Marketing Hub"},
        @{Title="Budget Tracker"; Type="List"; Url="/sites/finance/Lists/Budget"; HasUniquePermissions=$true; ParentUrl="https://contoso.sharepoint.com/sites/finance"; RoleAssignmentCount=2; SiteTitle="Finance Department"},
        @{Title="Team Calendar"; Type="List"; Url="/sites/teamsite/Lists/Calendar"; HasUniquePermissions=$false; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"},
        @{Title="Announcements"; Type="List"; Url="/sites/teamsite/Lists/Announcements"; HasUniquePermissions=$false; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"},
        @{Title="HR Policies"; Type="Document Library"; Url="/sites/hr/Policies"; HasUniquePermissions=$true; ParentUrl="https://contoso.sharepoint.com/sites/hr"; RoleAssignmentCount=3; SiteTitle="HR Portal"},
        @{Title="Style Library"; Type="Document Library"; Url="/sites/teamsite/Style Library"; HasUniquePermissions=$false; ParentUrl="https://contoso.sharepoint.com/sites/teamsite"; RoleAssignmentCount=0; SiteTitle="Team Collaboration Site"}
    )
    foreach ($item in $demoInheritance) { Add-SharePointInheritanceItem -InheritanceData $item }
    Write-ConsoleOutput "Added $($demoInheritance.Count) inheritance items"

    # Generate demo sharing links (includes anonymous edit links and many company-wide links)
    $demoLinks = @(
        # Anonymous links - triggers SHARE-001 (Critical - Anonymous Edit) and SHARE-002 (High - Anonymous links)
        @{GroupName="SharingLinks.abc125.AnonymousView.jkl012"; LinkType="Anonymous"; AccessLevel="View"; MemberCount=3; SiteTitle="Marketing Hub"; CreatedDate="2025-01-20"},
        @{GroupName="SharingLinks.abc128.AnonymousEdit.stu901"; LinkType="Anonymous"; AccessLevel="Edit"; MemberCount=1; SiteTitle="Team Collaboration Site"; CreatedDate="2025-03-01"},
        @{GroupName="SharingLinks.abc130.AnonymousEdit.xyz123"; LinkType="Anonymous"; AccessLevel="Edit"; MemberCount=2; SiteTitle="Marketing Hub"; CreatedDate="2024-12-15"},
        @{GroupName="SharingLinks.abc131.AnonymousView.abc789"; LinkType="Anonymous"; AccessLevel="View"; MemberCount=5; SiteTitle="Team Collaboration Site"; CreatedDate="2025-02-05"},
        @{GroupName="SharingLinks.abc132.AnonymousView.def456"; LinkType="Anonymous"; AccessLevel="View"; MemberCount=0; SiteTitle="HR Portal"; CreatedDate="2024-11-20"},

        # Company-wide links - triggers SHARE-003 (Medium - >10 company-wide links)
        @{GroupName="SharingLinks.abc123.OrganizationView.def456"; LinkType="Company-wide"; AccessLevel="View"; MemberCount=0; SiteTitle="Team Collaboration Site"; CreatedDate="2025-01-15"},
        @{GroupName="SharingLinks.abc124.OrganizationEdit.ghi789"; LinkType="Company-wide"; AccessLevel="Edit"; MemberCount=0; SiteTitle="Marketing Hub"; CreatedDate="2025-02-01"},
        @{GroupName="SharingLinks.abc133.OrganizationView.ghi012"; LinkType="Organization"; AccessLevel="View"; MemberCount=0; SiteTitle="HR Portal"; CreatedDate="2025-01-10"},
        @{GroupName="SharingLinks.abc134.OrganizationView.jkl345"; LinkType="Company-wide"; AccessLevel="View"; MemberCount=0; SiteTitle="Finance Department"; CreatedDate="2024-12-01"},
        @{GroupName="SharingLinks.abc135.OrganizationView.mno678"; LinkType="Organization"; AccessLevel="View"; MemberCount=0; SiteTitle="Executive Dashboard"; CreatedDate="2025-01-05"},
        @{GroupName="SharingLinks.abc136.OrganizationView.pqr901"; LinkType="Company-wide"; AccessLevel="View"; MemberCount=0; SiteTitle="Team Collaboration Site"; CreatedDate="2024-11-15"},
        @{GroupName="SharingLinks.abc137.OrganizationView.stu234"; LinkType="Organization"; AccessLevel="View"; MemberCount=0; SiteTitle="Marketing Hub"; CreatedDate="2025-02-10"},
        @{GroupName="SharingLinks.abc138.OrganizationView.vwx567"; LinkType="Company-wide"; AccessLevel="View"; MemberCount=0; SiteTitle="Team Collaboration Site"; CreatedDate="2024-10-20"},
        @{GroupName="SharingLinks.abc139.OrganizationView.yza890"; LinkType="Organization"; AccessLevel="View"; MemberCount=0; SiteTitle="HR Portal"; CreatedDate="2025-01-20"},
        @{GroupName="SharingLinks.abc140.OrganizationView.bcd123"; LinkType="Company-wide"; AccessLevel="View"; MemberCount=0; SiteTitle="Finance Department"; CreatedDate="2024-12-10"},
        @{GroupName="SharingLinks.abc141.OrganizationView.efg456"; LinkType="Organization"; AccessLevel="View"; MemberCount=0; SiteTitle="Marketing Hub"; CreatedDate="2025-02-05"},
        @{GroupName="SharingLinks.abc142.OrganizationView.hij789"; LinkType="Company-wide"; AccessLevel="View"; MemberCount=0; SiteTitle="Team Collaboration Site"; CreatedDate="2024-11-25"},

        # Specific people links (normal, lower risk)
        @{GroupName="SharingLinks.abc126.Flexible.mno345"; LinkType="Specific People"; AccessLevel="Edit"; MemberCount=5; SiteTitle="Team Collaboration Site"; CreatedDate="2025-02-10"},
        @{GroupName="SharingLinks.abc127.Flexible.pqr678"; LinkType="Specific People"; AccessLevel="View"; MemberCount=2; SiteTitle="HR Portal"; CreatedDate="2025-01-25"},
        @{GroupName="SharingLinks.abc129.Flexible.vwx234"; LinkType="Specific People"; AccessLevel="Edit"; MemberCount=8; SiteTitle="Finance Department"; CreatedDate="2025-02-15"}
    )
    foreach ($link in $demoLinks) { Add-SharePointSharingLink -LinkData $link }
    $anonymousCount = @($demoLinks | Where-Object {$_.LinkType -eq "Anonymous"}).Count
    $anonymousEditCount = @($demoLinks | Where-Object {$_.LinkType -eq "Anonymous" -and $_.AccessLevel -eq "Edit"}).Count
    $orgWideCount = @($demoLinks | Where-Object {$_.LinkType -in @("Company-wide", "Organization")}).Count
    Write-ConsoleOutput "Added $($demoLinks.Count) sharing links ($anonymousCount anonymous with $anonymousEditCount edit, $orgWideCount company-wide)"

    $metrics = Get-SharePointData -DataType "Metrics"
    Write-ConsoleOutput ""
    Write-ConsoleOutput "Demo mode ready!"
    Write-ConsoleOutput "Sites: $($metrics.TotalSites) | Users: $($metrics.TotalUsers) | Groups: $($metrics.TotalGroups)"
    Write-ConsoleOutput "External: $($metrics.ExternalUsers) | Roles: $($metrics.TotalRoleAssignments) | Inheritance Breaks: $($metrics.InheritanceBreaks) | Links: $($metrics.TotalSharingLinks)"
}

function Invoke-DemoEnrichment {
    <#
    .SYNOPSIS
    Simulates Graph enrichment for demo mode external users
    #>
    $users = Get-SharePointData -DataType "Users"
    $external = @($users | Where-Object { $_.Type -eq "External" -or $_.IsExternal })
    foreach ($u in $external) {
        $u.GraphUserType = "Guest"
        $u.GraphAccountEnabled = $true
        $u.GraphLastSignIn = (Get-Date).AddDays(-(Get-Random -Minimum 1 -Maximum 120)).ToString("o")
        $u.GraphCreatedDate = (Get-Date).AddDays(-(Get-Random -Minimum 30 -Maximum 365)).ToString("o")
        $u.GraphDisplayName = $u.Name
        $u.GraphEnriched = $true
    }
    return @{
        TotalExternal = $external.Count
        Enriched      = $external.Count
        Failed        = 0
    }
}

function Get-DemoPermissionsMatrix {
    <#
    .SYNOPSIS
    Generates demo permissions matrix data
    #>
    param(
        [string]$SiteUrl,
        [string]$ScanType = 'quick'
    )

    $siteName = if ($SiteUrl -match '/sites/(.+)$') { $matches[1] } else { 'Demo Site' }

    # Build a realistic permissions tree
    $tree = @(
        @{
            title = $siteName
            type = 'Site'
            url = $SiteUrl
            permissions = @(
                @{ principal = 'Site Owners'; role = 'Full Control' }
                @{ principal = 'Site Members'; role = 'Edit' }
                @{ principal = 'Site Visitors'; role = 'Read' }
            )
            children = @(
                @{
                    title = 'Documents'
                    type = 'Library'
                    url = "$SiteUrl/Shared Documents"
                    permissions = @()
                    children = @(
                        @{
                            title = 'Projects'
                            type = 'Folder'
                            url = "$SiteUrl/Shared Documents/Projects"
                            permissions = @(
                                @{ principal = 'Project Team'; role = 'Edit' }
                                @{ principal = 'external_consultant@partner.com'; role = 'Read' }
                            )
                        }
                        @{
                            title = 'Confidential Budget 2024.xlsx'
                            type = 'File'
                            url = "$SiteUrl/Shared Documents/Confidential Budget 2024.xlsx"
                            permissions = @(
                                @{ principal = 'Finance Team'; role = 'Full Control' }
                            )
                        }
                        @{
                            title = 'Team Handbook.docx'
                            type = 'File'
                            url = "$SiteUrl/Shared Documents/Team Handbook.docx"
                            permissions = @()
                        }
                    )
                }
                @{
                    title = 'HR Documents'
                    type = 'Library'
                    url = "$SiteUrl/HR Documents"
                    permissions = @(
                        @{ principal = 'HR Team'; role = 'Full Control' }
                        @{ principal = 'All Employees'; role = 'Read' }
                    )
                    children = @(
                        @{
                            title = 'Policies'
                            type = 'Folder'
                            url = "$SiteUrl/HR Documents/Policies"
                            permissions = @()
                        }
                        @{
                            title = 'Employee Contracts'
                            type = 'Folder'
                            url = "$SiteUrl/HR Documents/Employee Contracts"
                            permissions = @(
                                @{ principal = 'HR Managers'; role = 'Full Control' }
                            )
                        }
                    )
                }
                @{
                    title = 'Site Pages'
                    type = 'Library'
                    url = "$SiteUrl/SitePages"
                    permissions = @()
                    children = @(
                        @{
                            title = 'Home.aspx'
                            type = 'File'
                            url = "$SiteUrl/SitePages/Home.aspx"
                            permissions = @()
                        }
                    )
                }
                @{
                    title = 'Tasks'
                    type = 'List'
                    url = "$SiteUrl/Lists/Tasks"
                    permissions = @()
                    children = @()
                }
            )
        }
    )

    # Count totals
    $totalItems = 0
    $uniquePermissions = 0
    $principals = @{}

    function CountNode {
        param($node)
        $script:totalItems++
        if ($node.permissions -and $node.permissions.Count -gt 0) {
            $script:uniquePermissions++
            foreach ($perm in $node.permissions) {
                $principals[$perm.principal] = $true
            }
        }
        if ($node.children) {
            foreach ($child in $node.children) {
                CountNode -node $child
            }
        }
    }

    foreach ($node in $tree) {
        CountNode -node $node
    }

    return @{
        totalItems = $totalItems
        uniquePermissions = $uniquePermissions
        totalPrincipals = $principals.Count
        tree = $tree
        scanType = $ScanType
        scannedAt = (Get-Date).ToString("o")
    }
}
