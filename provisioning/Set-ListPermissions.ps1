<#
.SYNOPSIS
    Sets unique permissions on SharePoint Online lists for specified user groups.

.DESCRIPTION
    Configures list-level permission assignments by breaking permission inheritance
    and applying role assignments for specified SharePoint groups. Supports setting
    individual list permissions via parameters or bulk-applying permissions from a
    JSON configuration file.

    Requires the PnP.PowerShell module v2.x or later.

.PARAMETER SiteUrl
    The full URL of the target SharePoint site.

.PARAMETER ListName
    The display name of the target list. Required when not using -ConfigFile.

.PARAMETER GroupName
    The SharePoint group name to assign permissions to. Required when not using -ConfigFile.

.PARAMETER PermissionLevel
    The permission level to assign (e.g., Read, Contribute, Edit, Full Control).
    Required when not using -ConfigFile.

.PARAMETER ConfigFile
    Path to a JSON file containing bulk permission assignments. When specified,
    -ListName, -GroupName, and -PermissionLevel are ignored.

    JSON format:
    {
      "permissions": [
        {
          "listName": "Project Tracker",
          "assignments": [
            { "groupName": "Project Managers", "permissionLevel": "Edit" },
            { "groupName": "Stakeholders", "permissionLevel": "Read" }
          ]
        }
      ]
    }

.PARAMETER BreakInheritance
    Switch to break permission inheritance on the list before applying new
    permissions. If the list already has unique permissions, this has no effect.

.PARAMETER CopyExisting
    When used with -BreakInheritance, copies the existing role assignments before
    adding new ones. Default is $true.

.PARAMETER RemoveExistingAssignments
    Switch to remove all existing role assignments before applying new ones.
    Use with caution.

.PARAMETER Credential
    PSCredential object for authentication. If not provided, interactive login is used.

.PARAMETER WhatIf
    Shows what the script would do without making any changes.

.EXAMPLE
    .\Set-ListPermissions.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/workflow" `
        -ListName "Project Tracker" -GroupName "Project Managers" -PermissionLevel "Edit"

.EXAMPLE
    .\Set-ListPermissions.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/workflow" `
        -ListName "Project Tracker" -GroupName "Stakeholders" -PermissionLevel "Read" `
        -BreakInheritance

.EXAMPLE
    .\Set-ListPermissions.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/workflow" `
        -ConfigFile ".\permissions-config.json"

.EXAMPLE
    .\Set-ListPermissions.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/workflow" `
        -ConfigFile ".\permissions-config.json" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true, DefaultParameterSetName = 'Single')]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://[\w\-]+\.sharepoint\.com')]
    [string]$SiteUrl,

    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [ValidateNotNullOrEmpty()]
    [string]$ListName,

    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [ValidateNotNullOrEmpty()]
    [string]$GroupName,

    [Parameter(Mandatory = $true, ParameterSetName = 'Single')]
    [ValidateSet('Read', 'Contribute', 'Edit', 'Design', 'Full Control')]
    [string]$PermissionLevel,

    [Parameter(Mandatory = $true, ParameterSetName = 'Bulk')]
    [ValidateScript({ Test-Path $_ -PathType Leaf })]
    [string]$ConfigFile,

    [Parameter()]
    [switch]$BreakInheritance,

    [Parameter()]
    [bool]$CopyExisting = $true,

    [Parameter()]
    [switch]$RemoveExistingAssignments,

    [Parameter()]
    [System.Management.Automation.PSCredential]$Credential
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function Write-StatusLine {
    param([string]$Message, [string]$Status, [ConsoleColor]$Color = 'White')
    Write-Host ("{0,-60} " -f $Message) -NoNewline
    Write-Host "[$Status]" -ForegroundColor $Color
}

function Set-SingleListPermission {
    param(
        [string]$TargetListName,
        [string]$TargetGroupName,
        [string]$TargetPermissionLevel,
        [bool]$ShouldBreakInheritance,
        [bool]$ShouldCopyExisting,
        [bool]$ShouldRemoveExisting
    )

    $actionDescription = "Set '$TargetPermissionLevel' on list '$TargetListName' for group '$TargetGroupName'"

    if (-not $PSCmdlet.ShouldProcess($actionDescription, "Apply permission")) {
        Write-Host "  [WhatIf] $actionDescription" -ForegroundColor Yellow
        return
    }

    # Validate list exists
    Write-StatusLine "  Verifying list '$TargetListName'..." "RUNNING" Cyan
    try {
        $list = Get-PnPList -Identity $TargetListName -ErrorAction Stop
        Write-StatusLine "  List '$TargetListName' found" "PASS" Green
    }
    catch {
        Write-StatusLine "  List '$TargetListName' not found" "FAIL" Red
        Write-Warning "List '$TargetListName' does not exist on this site. Skipping."
        return
    }

    # Validate group exists
    Write-StatusLine "  Verifying group '$TargetGroupName'..." "RUNNING" Cyan
    try {
        $group = Get-PnPGroup -Identity $TargetGroupName -ErrorAction Stop
        Write-StatusLine "  Group '$TargetGroupName' found" "PASS" Green
    }
    catch {
        Write-StatusLine "  Group '$TargetGroupName' not found" "FAIL" Red
        Write-Warning "Group '$TargetGroupName' does not exist. Skipping."
        return
    }

    # Break inheritance if requested
    if ($ShouldBreakInheritance) {
        $hasUniquePerms = Get-PnPList -Identity $TargetListName | Select-Object -ExpandProperty HasUniqueRoleAssignments
        if (-not $hasUniquePerms) {
            Write-StatusLine "  Breaking permission inheritance..." "RUNNING" Cyan
            Set-PnPList -Identity $TargetListName -BreakRoleInheritance -CopyRoleAssignments:$ShouldCopyExisting
            Write-StatusLine "  Inheritance broken (copy existing: $ShouldCopyExisting)" "PASS" Green
        }
        else {
            Write-StatusLine "  List already has unique permissions" "SKIP" Yellow
        }
    }

    # Remove existing assignments if requested
    if ($ShouldRemoveExisting) {
        Write-StatusLine "  Removing existing role assignments..." "RUNNING" Cyan
        try {
            $roleAssignments = Get-PnPListRoleAssignment -Identity $TargetListName
            foreach ($assignment in $roleAssignments) {
                Remove-PnPListRoleAssignment -Identity $TargetListName -PrincipalId $assignment.PrincipalId -ErrorAction SilentlyContinue
            }
            Write-StatusLine "  Existing assignments removed" "PASS" Green
        }
        catch {
            Write-StatusLine "  Could not remove some assignments" "WARN" Yellow
            Write-Warning $_.Exception.Message
        }
    }

    # Apply the new permission
    Write-StatusLine "  Granting '$TargetPermissionLevel' to '$TargetGroupName'..." "RUNNING" Cyan
    try {
        Set-PnPListPermission -Identity $TargetListName -Group $TargetGroupName -AddRole $TargetPermissionLevel
        Write-StatusLine "  Permission applied: $TargetGroupName = $TargetPermissionLevel" "PASS" Green
    }
    catch {
        Write-StatusLine "  Failed to set permission" "FAIL" Red
        Write-Error "Failed to set '$TargetPermissionLevel' for '$TargetGroupName' on '$TargetListName': $($_.Exception.Message)"
    }
}

# ---------------------------------------------------------------------------
# Main
# ---------------------------------------------------------------------------
try {
    Write-Host "`n=== SharePoint List Permission Configuration ===" -ForegroundColor Cyan
    Write-Host "Site: $SiteUrl"
    Write-Host ("=" * 60)

    # Connect to SharePoint
    Write-StatusLine "Connecting to SharePoint Online..." "RUNNING" Cyan
    $connectParams = @{ Url = $SiteUrl }
    if ($Credential) {
        $connectParams['Credentials'] = $Credential
    }
    else {
        $connectParams['Interactive'] = $true
    }
    Connect-PnPOnline @connectParams
    Write-StatusLine "Connected to $SiteUrl" "PASS" Green

    if ($PSCmdlet.ParameterSetName -eq 'Bulk') {
        # --- Bulk mode from JSON config ---
        Write-Host "`nMode: Bulk (from config file)" -ForegroundColor Cyan
        Write-Host "Config: $ConfigFile"
        Write-Host ""

        $config = Get-Content -Path $ConfigFile -Raw | ConvertFrom-Json

        if (-not $config.permissions -or $config.permissions.Count -eq 0) {
            Write-Warning "Config file contains no permission entries."
            return
        }

        $totalAssignments = ($config.permissions | ForEach-Object { $_.assignments.Count } | Measure-Object -Sum).Sum
        Write-Host "Found $($config.permissions.Count) list(s) with $totalAssignments total assignment(s)`n"

        foreach ($listConfig in $config.permissions) {
            Write-Host "List: $($listConfig.listName)" -ForegroundColor White
            Write-Host ("-" * 40)

            foreach ($assignment in $listConfig.assignments) {
                Set-SingleListPermission `
                    -TargetListName $listConfig.listName `
                    -TargetGroupName $assignment.groupName `
                    -TargetPermissionLevel $assignment.permissionLevel `
                    -ShouldBreakInheritance $BreakInheritance.IsPresent `
                    -ShouldCopyExisting $CopyExisting `
                    -ShouldRemoveExisting $RemoveExistingAssignments.IsPresent
            }
            Write-Host ""
        }
    }
    else {
        # --- Single assignment mode ---
        Write-Host "`nMode: Single assignment" -ForegroundColor Cyan
        Write-Host "List : $ListName"
        Write-Host "Group: $GroupName"
        Write-Host "Level: $PermissionLevel"
        Write-Host ""

        Set-SingleListPermission `
            -TargetListName $ListName `
            -TargetGroupName $GroupName `
            -TargetPermissionLevel $PermissionLevel `
            -ShouldBreakInheritance $BreakInheritance.IsPresent `
            -ShouldCopyExisting $CopyExisting `
            -ShouldRemoveExisting $RemoveExistingAssignments.IsPresent
    }

    Write-Host ("=" * 60)
    Write-StatusLine "Permission configuration complete" "DONE" Green
    Write-Host ""
}
catch {
    Write-StatusLine "Permission configuration failed" "FAIL" Red
    Write-Error $_.Exception.Message
}
finally {
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
}
