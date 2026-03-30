<#
.SYNOPSIS
    Deploys the SharePoint Workflow Automation solution to a target SharePoint Online site.

.DESCRIPTION
    This script provisions the complete workflow automation solution including:
    - SharePoint site creation (if needed)
    - List templates from PnP provisioning XML
    - SPFx extension package deployment
    - List view and permission configuration

    Requires PnP.PowerShell module v2.x or later.

.PARAMETER SiteUrl
    The full URL of the target SharePoint site (e.g., https://contoso.sharepoint.com/sites/workflow-demo).

.PARAMETER TenantAdminUrl
    The SharePoint admin center URL (e.g., https://contoso-admin.sharepoint.com).
    Required only when creating a new site.

.PARAMETER Credential
    PSCredential object for authentication. If not provided, interactive login will be used.

.PARAMETER CreateSite
    Switch to create the target site if it does not exist.

.PARAMETER SiteTitle
    Title for the new site (used only with -CreateSite). Default: "Workflow Automation".

.PARAMETER SkipListTemplates
    Switch to skip provisioning list templates.

.PARAMETER SkipSPFx
    Switch to skip SPFx package deployment.

.PARAMETER WhatIf
    Shows what the script would do without making any changes.

.EXAMPLE
    .\Deploy-WorkflowSolution.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/workflow-demo"

.EXAMPLE
    .\Deploy-WorkflowSolution.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/workflow-demo" `
        -TenantAdminUrl "https://contoso-admin.sharepoint.com" -CreateSite

.EXAMPLE
    .\Deploy-WorkflowSolution.ps1 -SiteUrl "https://contoso.sharepoint.com/sites/workflow-demo" -WhatIf
#>

[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true)]
    [ValidatePattern('^https://[\w\-]+\.sharepoint\.com')]
    [string]$SiteUrl,

    [Parameter(Mandatory = $false)]
    [ValidatePattern('^https://[\w\-]+\-admin\.sharepoint\.com$')]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $false)]
    [PSCredential]$Credential,

    [switch]$CreateSite,

    [string]$SiteTitle = "Workflow Automation",

    [switch]$SkipListTemplates,

    [switch]$SkipSPFx
)

#region Configuration
$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$projectRoot = Split-Path -Parent $scriptRoot
$listTemplatesPath = Join-Path $projectRoot "list-templates"
$spfxPackagePath = Join-Path $projectRoot "spfx-extensions\sharepoint\solution\sharepoint-workflow-extensions.sppkg"
$logFile = Join-Path $scriptRoot "deploy-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

$listTemplates = @(
    @{ Name = "Project Tracker";               File = "project-tracker.xml" }
    @{ Name = "Document Approval Queue";        File = "document-approval.xml" }
    @{ Name = "Change Request Log";             File = "change-request.xml" }
    @{ Name = "Employee Onboarding Checklist";  File = "employee-onboarding.xml" }
)
#endregion

#region Logging
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
#endregion

#region Module Check
function Assert-PnPModule {
    Write-Log "Checking for PnP.PowerShell module..."

    if (-not (Get-Module -ListAvailable -Name PnP.PowerShell)) {
        Write-Log "PnP.PowerShell module not found. Installing..." "WARN"

        if ($PSCmdlet.ShouldProcess("PnP.PowerShell", "Install module")) {
            Install-Module -Name PnP.PowerShell -Scope CurrentUser -Force -AllowClobber
            Write-Log "PnP.PowerShell module installed successfully" "SUCCESS"
        }
    } else {
        $version = (Get-Module -ListAvailable -Name PnP.PowerShell |
                    Sort-Object Version -Descending |
                    Select-Object -First 1).Version
        Write-Log "PnP.PowerShell v$version found" "SUCCESS"
    }
}
#endregion

#region Connection
function Connect-ToSharePoint {
    param(
        [string]$Url
    )

    Write-Log "Connecting to SharePoint: $Url"

    $connectParams = @{ Url = $Url; Interactive = $true }

    if ($Credential) {
        $connectParams.Remove("Interactive")
        $connectParams["Credentials"] = $Credential
    }

    if ($PSCmdlet.ShouldProcess($Url, "Connect to SharePoint")) {
        Connect-PnPOnline @connectParams
        Write-Log "Connected to SharePoint successfully" "SUCCESS"
    }
}
#endregion

#region Site Creation
function New-WorkflowSite {
    Write-Log "Checking if site exists: $SiteUrl"

    if ($PSCmdlet.ShouldProcess($SiteUrl, "Create SharePoint site")) {
        try {
            # Connect to admin center to create site
            Connect-ToSharePoint -Url $TenantAdminUrl

            $existingSite = Get-PnPTenantSite -Url $SiteUrl -ErrorAction SilentlyContinue

            if ($existingSite) {
                Write-Log "Site already exists: $SiteUrl" "WARN"
                return
            }

            Write-Log "Creating new team site: $SiteTitle"

            New-PnPSite -Type TeamSite `
                -Title $SiteTitle `
                -Alias ($SiteUrl.Split('/')[-1]) `
                -Description "SharePoint Workflow Automation demo site"

            Write-Log "Site created: $SiteUrl" "SUCCESS"

            # Allow time for site provisioning
            Write-Log "Waiting for site provisioning to complete..."
            Start-Sleep -Seconds 30
        }
        catch {
            Write-Log "Failed to create site: $($_.Exception.Message)" "ERROR"
            throw
        }
    }
}
#endregion

#region List Provisioning
function Deploy-ListTemplates {
    Write-Log "Starting list template provisioning..."

    foreach ($template in $listTemplates) {
        $templateFile = Join-Path $listTemplatesPath $template.File

        if (-not (Test-Path $templateFile)) {
            Write-Log "Template file not found: $templateFile" "ERROR"
            continue
        }

        Write-Log "Provisioning: $($template.Name) from $($template.File)"

        if ($PSCmdlet.ShouldProcess($template.Name, "Apply PnP provisioning template")) {
            try {
                Invoke-PnPSiteTemplate -Path $templateFile -Handlers Lists, Fields, ContentTypes
                Write-Log "Provisioned: $($template.Name)" "SUCCESS"
            }
            catch {
                Write-Log "Failed to provision $($template.Name): $($_.Exception.Message)" "ERROR"
                Write-Log "Continuing with remaining templates..." "WARN"
            }
        }
    }
}
#endregion

#region SPFx Deployment
function Deploy-SPFxPackage {
    Write-Log "Starting SPFx package deployment..."

    if (-not (Test-Path $spfxPackagePath)) {
        Write-Log "SPFx package not found at: $spfxPackagePath" "ERROR"
        Write-Log "Run 'npm run package' in the spfx-extensions directory first" "WARN"
        return
    }

    if ($PSCmdlet.ShouldProcess("sharepoint-workflow-extensions.sppkg", "Deploy SPFx package")) {
        try {
            # Upload to app catalog
            Write-Log "Uploading SPFx package to app catalog..."
            $app = Add-PnPApp -Path $spfxPackagePath -Scope Tenant -Overwrite

            Write-Log "Package uploaded. App ID: $($app.Id)"

            # Deploy (make available)
            Write-Log "Publishing SPFx package..."
            Publish-PnPApp -Identity $app.Id -Scope Tenant

            # Install on the target site
            Write-Log "Installing SPFx package on site..."
            Install-PnPApp -Identity $app.Id -Scope Tenant

            Write-Log "SPFx package deployed and installed" "SUCCESS"
        }
        catch {
            Write-Log "Failed to deploy SPFx package: $($_.Exception.Message)" "ERROR"
            throw
        }
    }
}
#endregion

#region View Configuration
function Set-ListViewConfiguration {
    Write-Log "Configuring list views and field customizers..."

    if ($PSCmdlet.ShouldProcess("List views", "Configure custom formatting")) {
        try {
            # Register the Status field customizer on each list that has a Status field
            $statusLists = @("Project Tracker", "Document Approval Queue", "Change Request Log", "Employee Onboarding Checklist")
            $statusFieldCustomizerId = "a8e2c5d1-6b3f-4e7a-9c12-8d5f0e4b3a71"

            foreach ($listTitle in $statusLists) {
                Write-Log "Configuring field customizer on: $listTitle"

                try {
                    $list = Get-PnPList -Identity $listTitle -ErrorAction SilentlyContinue
                    if ($null -eq $list) {
                        Write-Log "List not found: $listTitle" "WARN"
                        continue
                    }

                    # Get the appropriate status field for each list
                    $statusFieldNames = @{
                        "Project Tracker"                 = "ProjectStatus"
                        "Document Approval Queue"         = "ApprovalStatus"
                        "Change Request Log"              = "CRStatus"
                        "Employee Onboarding Checklist"   = "OBStatus"
                    }

                    $fieldName = $statusFieldNames[$listTitle]
                    $field = Get-PnPField -List $listTitle -Identity $fieldName -ErrorAction SilentlyContinue

                    if ($field) {
                        Set-PnPField -List $listTitle -Identity $fieldName -Values @{
                            ClientSideComponentId   = $statusFieldCustomizerId
                            ClientSideComponentProperties = "{}"
                        }
                        Write-Log "Field customizer registered on $listTitle.$fieldName" "SUCCESS"
                    }
                }
                catch {
                    Write-Log "Error configuring $listTitle : $($_.Exception.Message)" "WARN"
                }
            }

            # Register command set on Project Tracker
            Write-Log "Registering Bulk Actions command set on Project Tracker..."
            $commandSetId = "bf1e4ec6-92a7-4c3d-b579-dfb4ce402707"

            Add-PnPCustomAction `
                -Name "BulkActionsCommandSet" `
                -Title "Bulk Actions" `
                -Location "ClientSideExtension.ListViewCommandSet.CommandBar" `
                -ClientSideComponentId $commandSetId `
                -ClientSideComponentProperties '{"statusFieldName":"ProjectStatus"}' `
                -RegistrationId 100 `
                -RegistrationType List `
                -Scope Web

            Write-Log "Command set registered" "SUCCESS"
        }
        catch {
            Write-Log "Failed to configure views: $($_.Exception.Message)" "ERROR"
        }
    }
}
#endregion

#region Main Execution
function Invoke-Deployment {
    Write-Log "========================================="
    Write-Log "SharePoint Workflow Automation Deployment"
    Write-Log "========================================="
    Write-Log "Site URL:         $SiteUrl"
    Write-Log "Tenant Admin URL: $TenantAdminUrl"
    Write-Log "Create Site:      $CreateSite"
    Write-Log "Skip Templates:   $SkipListTemplates"
    Write-Log "Skip SPFx:        $SkipSPFx"
    Write-Log "WhatIf:           $WhatIfPreference"
    Write-Log "Log file:         $logFile"
    Write-Log "========================================="

    $stopwatch = [System.Diagnostics.Stopwatch]::StartNew()

    try {
        # Step 1: Check prerequisites
        Assert-PnPModule

        # Step 2: Create site (if requested)
        if ($CreateSite) {
            if (-not $TenantAdminUrl) {
                Write-Log "TenantAdminUrl is required when using -CreateSite" "ERROR"
                return
            }
            New-WorkflowSite
        }

        # Step 3: Connect to target site
        Connect-ToSharePoint -Url $SiteUrl

        # Step 4: Deploy list templates
        if (-not $SkipListTemplates) {
            Deploy-ListTemplates
        } else {
            Write-Log "Skipping list template provisioning (SkipListTemplates flag set)" "WARN"
        }

        # Step 5: Deploy SPFx package
        if (-not $SkipSPFx) {
            Deploy-SPFxPackage
        } else {
            Write-Log "Skipping SPFx deployment (SkipSPFx flag set)" "WARN"
        }

        # Step 6: Configure views and extensions
        if (-not $SkipListTemplates -and -not $SkipSPFx) {
            Set-ListViewConfiguration
        }

        $stopwatch.Stop()
        Write-Log "========================================="
        Write-Log "Deployment completed in $($stopwatch.Elapsed.ToString('mm\:ss'))" "SUCCESS"
        Write-Log "========================================="
        Write-Log ""
        Write-Log "Next steps:"
        Write-Log "  1. Import Power Automate flows from power-automate-flows/"
        Write-Log "  2. Update flow parameters with your list GUIDs"
        Write-Log "  3. Test the solution at: $SiteUrl"
    }
    catch {
        $stopwatch.Stop()
        Write-Log "========================================="
        Write-Log "Deployment FAILED after $($stopwatch.Elapsed.ToString('mm\:ss'))" "ERROR"
        Write-Log "Error: $($_.Exception.Message)" "ERROR"
        Write-Log "========================================="
        throw
    }
    finally {
        Disconnect-PnPOnline -ErrorAction SilentlyContinue
    }
}

# Execute
Invoke-Deployment
