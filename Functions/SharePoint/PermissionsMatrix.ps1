function Get-SitePermissionsMatrix {
    <#
    .SYNOPSIS
        Collects permissions matrix for a SharePoint site including files and folders
    .DESCRIPTION
        Recursively traverses a SharePoint site and collects role assignments at every level:
        Site → Lists/Libraries → Folders → Files
    .PARAMETER SiteUrl
        The URL of the SharePoint site to scan
    .PARAMETER ScanType
        'quick' - Only items with unique permissions (faster)
        'full' - All files and folders (comprehensive but slower)
    #>
    param(
        [Parameter(Mandatory=$true)]
        [string]$SiteUrl,

        [Parameter(Mandatory=$false)]
        [ValidateSet('quick', 'full')]
        [string]$ScanType = 'quick'
    )

    $totalItems = 0
    $uniquePermissions = 0
    $principals = @{}
    $tree = [System.Collections.ArrayList]::new()

    try {
        Write-ActivityLog "Starting permissions matrix scan for $SiteUrl (type: $ScanType)"

        # Get site permissions
        $site = Get-PnPWeb
        $siteNode = @{
            title = $site.Title
            type = 'Site'
            url = $site.Url
            permissions = [System.Collections.ArrayList]::new()
            children = [System.Collections.ArrayList]::new()
        }

        # Get site role assignments
        # Get-PnPRoleAssignment does not exist in PnP.PowerShell 3.x; use Get-PnPProperty instead.
        $siteRoles = Get-PnPProperty -ClientObject $site -Property RoleAssignments -ErrorAction SilentlyContinue
        if ($siteRoles) {
            foreach ($role in $siteRoles) {
                Get-PnPProperty -ClientObject $role -Property Member, RoleDefinitionBindings -ErrorAction SilentlyContinue
                $principal = $role.Member.Title
                $roleBinding = $role.RoleDefinitionBindings[0].Name

                $siteNode.permissions += @{
                    principal = $principal
                    role = $roleBinding
                }
                $principals[$principal] = $true
                $uniquePermissions++
            }
        }
        $totalItems++

        # Get lists/libraries
        # -Includes ensures HasUniqueRoleAssignments is loaded; accessing an unloaded CSOM
        # property throws PropertyOrFieldNotInitializedException which silently empties children.
        $lists = Get-PnPList -Includes HasUniqueRoleAssignments | Where-Object { -not $_.Hidden -and $_.ItemCount -gt 0 }

        foreach ($list in $lists) {
            $listNode = @{
                title = $list.Title
                type = if ($list.BaseTemplate -eq 101) { 'Library' } else { 'List' }
                url = "$($site.Url)/$($list.RootFolder.ServerRelativeUrl)"
                permissions = [System.Collections.ArrayList]::new()
                children = [System.Collections.ArrayList]::new()
            }

            # Check list permissions
            if ($list.HasUniqueRoleAssignments) {
                try {
                    $listRoles = Get-PnPProperty -ClientObject $list -Property RoleAssignments -ErrorAction SilentlyContinue
                    if ($listRoles) {
                        foreach ($role in $listRoles) {
                            Get-PnPProperty -ClientObject $role -Property Member, RoleDefinitionBindings -ErrorAction SilentlyContinue
                            $principal = $role.Member.Title
                            $roleBinding = $role.RoleDefinitionBindings[0].Name

                            $listNode.permissions += @{
                                principal = $principal
                                role = $roleBinding
                            }
                            $principals[$principal] = $true
                            $uniquePermissions++
                        }
                    }
                } catch {
                    Write-ActivityLog "Warning: Could not get permissions for list $($list.Title): $($_.Exception.Message)"
                }
            }
            $totalItems++

            # Get folders and files (if full scan OR if quick scan and list has unique permissions)
            if ($ScanType -eq 'full' -or ($ScanType -eq 'quick' -and $list.HasUniqueRoleAssignments)) {
                try {
                    # Get all items in the list
                    # -Includes loads HasUniqueRoleAssignments so the hierarchy check doesn't throw.
                    $items = Get-PnPListItem -List $list.Title -PageSize 500 -Includes "HasUniqueRoleAssignments" -ErrorAction SilentlyContinue

                    foreach ($item in $items) {
                        # Skip if quick scan and item inherits permissions
                        if ($ScanType -eq 'quick' -and -not $item.HasUniqueRoleAssignments) {
                            continue
                        }

                        $itemType = if ($item.FileSystemObjectType -eq 'Folder') { 'Folder' } else { 'File' }
                        $itemNode = @{
                            title = $item.FieldValues.FileLeafRef
                            type = $itemType
                            url = "$($site.Url)$($item.FieldValues.FileRef)"
                            permissions = @()
                        }

                        if ($item.HasUniqueRoleAssignments) {
                            try {
                                $itemRoles = Get-PnPProperty -ClientObject $item -Property RoleAssignments -ErrorAction SilentlyContinue
                                if ($itemRoles) {
                                    foreach ($role in $itemRoles) {
                                        Get-PnPProperty -ClientObject $role -Property Member, RoleDefinitionBindings -ErrorAction SilentlyContinue
                                        $principal = $role.Member.Title
                                        $roleBinding = $role.RoleDefinitionBindings[0].Name

                                        $itemNode.permissions += @{
                                            principal = $principal
                                            role = $roleBinding
                                        }
                                        $principals[$principal] = $true
                                        $uniquePermissions++
                                    }
                                }
                            } catch {
                                Write-ActivityLog "Warning: Could not get permissions for item $($item.FieldValues.FileLeafRef): $($_.Exception.Message)"
                            }
                        }

                        [void]$listNode.children.Add($itemNode)
                        $totalItems++
                    }
                } catch {
                    Write-ActivityLog "Warning: Could not scan items in list $($list.Title): $($_.Exception.Message)"
                }
            }

            [void]$siteNode.children.Add($listNode)
        }

        [void]$tree.Add($siteNode)

        Write-ActivityLog "Permissions matrix scan complete: $totalItems items, $uniquePermissions unique permissions"

        return @{
            totalItems = $totalItems
            uniquePermissions = $uniquePermissions
            totalPrincipals = $principals.Count
            tree = $tree
            scanType = $ScanType
            scannedAt = (Get-Date).ToString("o")
        }

    } catch {
        Write-ActivityLog "ERROR: Permissions matrix scan failed: $($_.Exception.Message)"
        throw
    }
}
