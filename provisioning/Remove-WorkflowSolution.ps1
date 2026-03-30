<#
.SYNOPSIS
    Removes the SharePoint Workflow Automation solution from a target site.

.DESCRIPTION
    Cleanup script for development and test environments. Removes:
    - Custom actions (SPFx command sets, field customizers)
    - SharePoint lists created by the solution
    - SPFx app package from the site
    - Optionally deletes the entire site

    WARNING: This operation is destructive and cannot be undone.
    Always run with -WhatIf first to review planned changes.

.PARAMETER SiteUrl
    The full URL of the SharePoint site to clean up.

.PARAMETER Credential
    PSCredential object for authentication.

.PARAMETER RemoveLists
    Switch to remove the provisioned lists. Default: $false.

.PARAMETER RemoveApp
    Switch to uninstall and remove the SPFx app. Default: $false.

.PARAMETER RemoveSite
    Switch to delete the entire site. Requires TenantAdminUrl. Default: $false.

.PARAMETER TenantAdminUrl
    Admin center URL, required only when using -RemoveSite.

.PARAMETER Force
    Suppresses confirmation prompts.

.EXAMPLE
    .\Remove-WorkflowSolution.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/workflow-demo" -WhatIf

.EXAMPLE
    .\Remove-WorkflowSolution.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/workflow-demo" `
        -RemoveLists -RemoveApp -Force
#>

[CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://[\w\-]+\.sharepoint\.com')]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [PSCredential]$Credential,

    [switch]$RemoveLists,

    [switch]$RemoveApp,

    [switch]$RemoveSite,

    [Parameter(Mandatory = $false)]
    [string]$TenantAdminUrl,

    [switch]$Force
)

$ErrorActionPreference = "Stop"
$logFile = Join-Path (Split-Path -Parent $MyInvocation.MyCommand.Definition) "remove-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

function Write-Log {
    param(
        [string]$Message,
        [ValidateSet("INFO", "WARN", "ERROR", "SUCCESS")]
        [string]$Level = "INFO"
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"

    switch ($Level) {
        "INFO"    { Write-Host $logEntry -ForegroundColor Cyan }
        "WARN"    { Write-Host $logEntry -ForegroundColor Yellow }
        "ERROR"   { Write-Host $logEntry -ForegroundColor Red }
        "SUCCESS" { Write-Host $logEntry -ForegroundColor Green }
    }

    Add-Content -Path $logFile -Value $logEntry
}

$listsToRemove = @(
    "Project Tracker",
    "Document Approval Queue",
    "Change Request Log",
    "Employee Onboarding Checklist"
)

$appProductId = "c4d92f1a-8e56-4b3c-a7d1-9f0e2c5b8a34"

try {
    Write-Log "============================================"
    Write-Log "SharePoint Workflow Automation - Cleanup"
    Write-Log "============================================"
    Write-Log "Site URL:      $SiteUrl"
    Write-Log "Remove Lists:  $RemoveLists"
    Write-Log "Remove App:    $RemoveApp"
    Write-Log "Remove Site:   $RemoveSite"
    Write-Log "============================================"

    if (-not $Force -and -not $WhatIfPreference) {
        $confirmation = Read-Host "This will remove solution components from $SiteUrl. Continue? (y/N)"
        if ($confirmation -ne 'y') {
            Write-Log "Operation cancelled by user" "WARN"
            return
        }
    }

    # Connect
    $connectParams = @{ Url = $SiteUrl; Interactive = $true }
    if ($Credential) {
        $connectParams.Remove("Interactive")
        $connectParams["Credentials"] = $Credential
    }

    if ($PSCmdlet.ShouldProcess($SiteUrl, "Connect to SharePoint")) {
        Connect-PnPOnline @connectParams
        Write-Log "Connected to SharePoint" "SUCCESS"
    }

    # Step 1: Remove custom actions
    Write-Log "Removing custom actions..."
    if ($PSCmdlet.ShouldProcess("Custom actions", "Remove SPFx registrations")) {
        $customActions = Get-PnPCustomAction -Scope Web
        foreach ($action in $customActions) {
            if ($action.Name -like "BulkActions*" -or
                $action.ClientSideComponentId -eq "bf1e4ec6-92a7-4c3d-b579-dfb4ce402707" -or
                $action.ClientSideComponentId -eq "a8e2c5d1-6b3f-4e7a-9c12-8d5f0e4b3a71") {
                Write-Log "  Removing custom action: $($action.Name) ($($action.Id))"
                Remove-PnPCustomAction -Identity $action.Id -Scope Web -Force
                Write-Log "  Removed: $($action.Name)" "SUCCESS"
            }
        }
    }

    # Step 2: Remove lists
    if ($RemoveLists) {
        Write-Log "Removing provisioned lists..."
        foreach ($listTitle in $listsToRemove) {
            if ($PSCmdlet.ShouldProcess($listTitle, "Remove list")) {
                try {
                    $list = Get-PnPList -Identity $listTitle -ErrorAction SilentlyContinue
                    if ($list) {
                        Write-Log "  Removing list: $listTitle"
                        Remove-PnPList -Identity $listTitle -Force
                        Write-Log "  Removed: $listTitle" "SUCCESS"
                    } else {
                        Write-Log "  List not found (may already be removed): $listTitle" "WARN"
                    }
                }
                catch {
                    Write-Log "  Error removing $listTitle : $($_.Exception.Message)" "ERROR"
                }
            }
        }
    }

    # Step 3: Remove SPFx app
    if ($RemoveApp) {
        Write-Log "Removing SPFx app package..."
        if ($PSCmdlet.ShouldProcess("sharepoint-workflow-extensions", "Uninstall and remove app")) {
            try {
                $app = Get-PnPApp -Scope Tenant | Where-Object { $_.ProductId -eq $appProductId }
                if ($app) {
                    Write-Log "  Uninstalling app: $($app.Title)"
                    Uninstall-PnPApp -Identity $app.Id -Scope Tenant
                    Start-Sleep -Seconds 5

                    Write-Log "  Removing app from catalog"
                    Remove-PnPApp -Identity $app.Id -Scope Tenant
                    Write-Log "  App removed" "SUCCESS"
                } else {
                    Write-Log "  App not found in tenant catalog" "WARN"
                }
            }
            catch {
                Write-Log "  Error removing app: $($_.Exception.Message)" "ERROR"
            }
        }
    }

    # Step 4: Remove content types and site columns
    Write-Log "Removing content types and site columns..."
    if ($PSCmdlet.ShouldProcess("Site columns", "Remove workflow automation fields")) {
        $workflowFields = Get-PnPField | Where-Object { $_.Group -eq "Workflow Automation" }
        foreach ($field in $workflowFields) {
            try {
                Write-Log "  Removing field: $($field.InternalName)"
                Remove-PnPField -Identity $field.Id -Force
            }
            catch {
                Write-Log "  Could not remove field $($field.InternalName): $($_.Exception.Message)" "WARN"
            }
        }

        $contentTypeIds = @(
            "0x01004A8C3E7B2D5F6910AB3C4D5E6F7890",
            "0x01005B9D4E8C3A2F7160BC5D6E7F8901",
            "0x01006C0E5F9D8A3B7241CD6E7F8A9B01",
            "0x01007D1E0F2A3B4C5D6E7F8A9B0C1D2E"
        )
        foreach ($ctId in $contentTypeIds) {
            try {
                $ct = Get-PnPContentType | Where-Object { $_.StringId -like "$ctId*" }
                if ($ct) {
                    Write-Log "  Removing content type: $($ct.Name)"
                    Remove-PnPContentType -Identity $ct.Name -Force
                }
            }
            catch {
                Write-Log "  Could not remove content type $ctId : $($_.Exception.Message)" "WARN"
            }
        }
    }

    # Step 5: Remove site (if requested)
    if ($RemoveSite) {
        if (-not $TenantAdminUrl) {
            Write-Log "TenantAdminUrl is required to remove the site" "ERROR"
        } else {
            if ($PSCmdlet.ShouldProcess($SiteUrl, "Delete SharePoint site")) {
                Disconnect-PnPOnline -ErrorAction SilentlyContinue
                Connect-PnPOnline -Url $TenantAdminUrl -Interactive

                Write-Log "Deleting site: $SiteUrl"
                Remove-PnPTenantSite -Url $SiteUrl -Force
                Write-Log "Site deleted" "SUCCESS"

                Write-Log "Removing from recycle bin..."
                Remove-PnPTenantDeletedSite -Url $SiteUrl -Force
                Write-Log "Site permanently removed" "SUCCESS"
            }
        }
    }

    Write-Log "============================================"
    Write-Log "Cleanup completed" "SUCCESS"
    Write-Log "============================================"
}
catch {
    Write-Log "============================================"
    Write-Log "Cleanup FAILED: $($_.Exception.Message)" "ERROR"
    Write-Log "============================================"
    throw
}
finally {
    Disconnect-PnPOnline -ErrorAction SilentlyContinue
}
