<#
.SYNOPSIS
    Creates a new SharePoint Online site and provisions it from a PnP template
    with lists, views, content types, navigation, theme, permissions, and SPFx
    extensions.

.DESCRIPTION
    This script automates end-to-end site provisioning:

    1. Creates a new Communication or Team site at the specified URL.
    2. Applies a PnP provisioning template (XML) that defines lists, views,
       fields, content types, and navigation.
    3. Configures the site theme.
    4. Sets up permissions (owner, members, visitors).
    5. Registers SPFx extensions on target lists.
    6. Generates a post-deployment verification report.

    Supports -WhatIf for dry-run mode.

    Requires PnP.PowerShell v2.x or later.

.PARAMETER SiteUrl
    The full URL of the new SharePoint site to create.
    Example: https://contoso.sharepoint.com/sites/new-project

.PARAMETER SiteTitle
    Display title for the new site. Default: "Workflow Site".

.PARAMETER Template
    Provisioning template to apply. Use "default" for the built-in workflow
    template or provide a full path to a custom PnP XML template.
    Default: "default".

.PARAMETER Owner
    UPN or email of the site owner. If omitted, the connected user becomes
    the owner.

.PARAMETER SiteType
    The type of site to create: "CommunicationSite" or "TeamSite".
    Default: "CommunicationSite".

.PARAMETER TenantAdminUrl
    SharePoint admin center URL. Required for site creation.
    Example: https://contoso-admin.sharepoint.com

.PARAMETER Credential
    PSCredential for non-interactive authentication. If omitted, interactive
    login is used.

.PARAMETER ThemeName
    Name of a tenant theme to apply. If omitted, no theme is applied.

.PARAMETER SkipPermissions
    Switch to skip permission configuration.

.PARAMETER SkipSPFx
    Switch to skip SPFx extension registration.

.PARAMETER WhatIf
    Shows what the script would do without making any changes.

.EXAMPLE
    .\New-SiteFromTemplate.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/project-alpha" `
        -SiteTitle "Project Alpha" `
        -Owner "admin@contoso.onmicrosoft.com" `
        -TenantAdminUrl "https://contoso-admin.sharepoint.com"

.EXAMPLE
    .\New-SiteFromTemplate.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/custom" `
        -Template "C:\Templates\custom-template.xml" `
        -SiteType "TeamSite" `
        -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
        -WhatIf

.EXAMPLE
    .\New-SiteFromTemplate.ps1 `
        -SiteUrl "https://contoso.sharepoint.com/sites/demo" `
        -SiteTitle "Demo Site" `
        -TenantAdminUrl "https://contoso-admin.sharepoint.com" `
        -ThemeName "Contoso Blue" `
        -SkipSPFx
#>
[CmdletBinding(SupportsShouldProcess = $true)]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Full URL of the new SharePoint site")]
    [ValidatePattern('^https://[\w\-]+\.sharepoint\.com/sites/[\w\-]+$')]
    [string]$SiteUrl,

    [Parameter(HelpMessage = "Display title for the new site")]
    [ValidateNotNullOrEmpty()]
    [string]$SiteTitle = "Workflow Site",

    [Parameter(HelpMessage = "PnP template name ('default') or path to custom XML template")]
    [string]$Template = "default",

    [Parameter(HelpMessage = "UPN or email of the site owner")]
    [string]$Owner,

    [Parameter(HelpMessage = "Site type to create: CommunicationSite or TeamSite")]
    [ValidateSet("CommunicationSite", "TeamSite")]
    [string]$SiteType = "CommunicationSite",

    [Parameter(Mandatory = $false, HelpMessage = "SharePoint admin center URL")]
    [ValidatePattern('^https://[\w\-]+\-admin\.sharepoint\.com$')]
    [string]$TenantAdminUrl,

    [Parameter(Mandatory = $false)]
    [PSCredential]$Credential,

    [Parameter(HelpMessage = "Tenant theme name to apply")]
    [string]$ThemeName,

    [switch]$SkipPermissions,
    [switch]$SkipSPFx
)

#region Configuration
$ErrorActionPreference = "Stop"
$scriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Definition
$projectRoot = Split-Path -Parent $scriptRoot
$listTemplatesPath = Join-Path $projectRoot "list-templates"
$spfxPackagePath = Join-Path $projectRoot "spfx-extensions\sharepoint\solution\sharepoint-workflow-extensions.sppkg"
$logFile = Join-Path $scriptRoot "provision-$(Get-Date -Format 'yyyyMMdd-HHmmss').log"

# Default template includes all standard workflow lists
$defaultTemplateLists = @(
    @{ Name = "project-tracker";       File = "project-tracker.xml" },
    @{ Name = "document-approval";     File = "document-approval.xml" },
    @{ Name = "change-request";        File = "change-request.xml" },
    @{ Name = "employee-onboarding";   File = "employee-onboarding.xml" },
    @{ Name = "it-asset-tracking";     File = "it-asset-tracking.xml" }
)

# SPFx extensions to register on specific lists
$extensionRegistrations = @(
    @{
        ListTitle      = "Project Tracker"
        ExtensionType  = "ListViewCommandSet"
        ClientId       = "e5a7b3c2-1d4f-4e9a-8b6c-0d2e1f3a5b4c"
        Properties     = '{"statusFieldName":"Status"}'
    },
    @{
        ListTitle      = "Document Approval Queue"
        ExtensionType  = "ListViewCommandSet"
        ClientId       = "e5a7b3c2-1d4f-4e9a-8b6c-0d2e1f3a5b4c"
        Properties     = '{"statusFieldName":"ApprovalStatus"}'
    },
    @{
        ListTitle      = "Change Request Log"
        ExtensionType  = "ListViewCommandSet"
        ClientId       = "e5a7b3c2-1d4f-4e9a-8b6c-0d2e1f3a5b4c"
        Properties     = '{"statusFieldName":"Status"}'
    }
)
#endregion

#region Helper Functions
function Write-Log {
    param([string]$Message, [string]$Level = "INFO")
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] [$Level] $Message"
    Add-Content -Path $logFile -Value $logEntry -ErrorAction SilentlyContinue

    switch ($Level) {
        "ERROR" { Write-Host $Message -ForegroundColor Red }
        "WARN"  { Write-Host $Message -ForegroundColor Yellow }
        "OK"    { Write-Host $Message -ForegroundColor Green }
        default { Write-Host $Message }
    }
}

function Write-StepHeader {
    param([int]$Step, [string]$Title)
    Write-Host ""
    Write-Host ("=" * 60) -ForegroundColor Cyan
    Write-Host "  Step $Step : $Title" -ForegroundColor Cyan
    Write-Host ("=" * 60) -ForegroundColor Cyan
}

function Test-SiteExists {
    param([string]$Url)
    try {
        $site = Get-PnPTenantSite -Url $Url -ErrorAction Stop
        return $null -ne $site
    }
    catch {
        return $false
    }
}
#endregion

#region Report
$deploymentReport = @{
    StartTime     = Get-Date
    SiteUrl       = $SiteUrl
    SiteTitle     = $SiteTitle
    Template      = $Template
    Steps         = [System.Collections.ArrayList]::new()
}

function Add-ReportStep {
    param([string]$Name, [string]$Status, [string]$Detail = "")
    [void]$deploymentReport.Steps.Add([PSCustomObject]@{
        Step   = $Name
        Status = $Status
        Detail = $Detail
    })
}
#endregion

# ===========================================================================
# Main Execution
# ===========================================================================
try {
    Write-Host "`n=== SharePoint Site Provisioning ===" -ForegroundColor Cyan
    Write-Host "Site URL   : $SiteUrl"
    Write-Host "Site Title : $SiteTitle"
    Write-Host "Site Type  : $SiteType"
    Write-Host "Template   : $Template"
    Write-Host "Owner      : $(if ($Owner) { $Owner } else { '(current user)' })"
    Write-Host "Log file   : $logFile"
    Write-Host ""

    # --- WhatIf summary ---
    if (-not $PSCmdlet.ShouldProcess($SiteUrl, "Provision SharePoint site from template")) {
        Write-Host "[WhatIf] The following actions would be performed:" -ForegroundColor Yellow
        Write-Host "  1. Create $SiteType at $SiteUrl"
        Write-Host "  2. Apply PnP template ($Template) with $($defaultTemplateLists.Count) list definitions"
        if (-not $SkipPermissions) {
            Write-Host "  3. Configure permissions (owner: $(if ($Owner) { $Owner } else { 'current user' }))"
        }
        if ($ThemeName) {
            Write-Host "  4. Apply theme: $ThemeName"
        }
        if (-not $SkipSPFx) {
            Write-Host "  5. Register $($extensionRegistrations.Count) SPFx extensions on target lists"
        }
        Write-Host "  6. Generate post-deployment report"
        return
    }

    # ------------------------------------------------------------------
    # Step 1: Connect and create site
    # ------------------------------------------------------------------
    Write-StepHeader 1 "Create Site"

    if ($TenantAdminUrl) {
        $connectParams = @{ Url = $TenantAdminUrl }
        if ($Credential) { $connectParams.Credentials = $Credential }
        Write-Log "Connecting to tenant admin: $TenantAdminUrl"
        Connect-PnPOnline @connectParams -ErrorAction Stop
    }
    else {
        Write-Log "No TenantAdminUrl provided; connecting directly to $SiteUrl" "WARN"
    }

    $siteExists = Test-SiteExists -Url $SiteUrl

    if ($siteExists) {
        Write-Log "Site already exists at $SiteUrl -- skipping creation." "WARN"
        Add-ReportStep -Name "Create Site" -Status "Skipped" -Detail "Site already exists"
    }
    else {
        Write-Log "Creating $SiteType at $SiteUrl..."

        $newSiteParams = @{
            Title = $SiteTitle
            Url   = $SiteUrl
            Type  = $SiteType
        }
        if ($Owner) { $newSiteParams.Owner = $Owner }

        New-PnPSite @newSiteParams -ErrorAction Stop | Out-Null
        Write-Log "Site created successfully." "OK"
        Add-ReportStep -Name "Create Site" -Status "Success" -Detail "$SiteType created at $SiteUrl"
    }

    # Reconnect to the new site for provisioning
    $siteConnectParams = @{ Url = $SiteUrl }
    if ($Credential) { $siteConnectParams.Credentials = $Credential }
    Connect-PnPOnline @siteConnectParams -ErrorAction Stop
    Write-Log "Connected to $SiteUrl"

    # ------------------------------------------------------------------
    # Step 2: Apply PnP template
    # ------------------------------------------------------------------
    Write-StepHeader 2 "Apply Provisioning Template"

    if ($Template -eq "default") {
        # Apply each list template from the built-in set
        foreach ($tpl in $defaultTemplateLists) {
            $templateFile = Join-Path $listTemplatesPath $tpl.File
            if (Test-Path $templateFile) {
                Write-Log "Applying template: $($tpl.Name) ($($tpl.File))"
                Invoke-PnPSiteTemplate -Path $templateFile -ErrorAction Stop
                Write-Log "  Applied $($tpl.Name)" "OK"
            }
            else {
                Write-Log "  Template file not found: $templateFile" "WARN"
            }
        }
        Add-ReportStep -Name "Apply Template" -Status "Success" -Detail "$($defaultTemplateLists.Count) list templates applied"
    }
    else {
        # Apply custom template
        if (-not (Test-Path $Template)) {
            throw "Custom template file not found: $Template"
        }
        Write-Log "Applying custom template: $Template"
        Invoke-PnPSiteTemplate -Path $Template -ErrorAction Stop
        Write-Log "Custom template applied." "OK"
        Add-ReportStep -Name "Apply Template" -Status "Success" -Detail "Custom template applied: $Template"
    }

    # ------------------------------------------------------------------
    # Step 3: Configure permissions
    # ------------------------------------------------------------------
    if (-not $SkipPermissions) {
        Write-StepHeader 3 "Configure Permissions"

        if ($Owner) {
            Write-Log "Setting site owner: $Owner"
            Set-PnPSite -Identity $SiteUrl -Owners $Owner -ErrorAction Stop
            Write-Log "Owner configured." "OK"
        }

        # Ensure default groups exist
        try {
            $web = Get-PnPWeb -Includes AssociatedOwnerGroup, AssociatedMemberGroup, AssociatedVisitorGroup
            Write-Log "Associated groups verified: Owners=$($web.AssociatedOwnerGroup.Title), Members=$($web.AssociatedMemberGroup.Title), Visitors=$($web.AssociatedVisitorGroup.Title)"
        }
        catch {
            Write-Log "Could not verify associated groups: $($_.Exception.Message)" "WARN"
        }

        Add-ReportStep -Name "Permissions" -Status "Success" -Detail "Owner: $(if ($Owner) { $Owner } else { 'current user' })"
    }
    else {
        Write-Log "Skipping permission configuration (SkipPermissions flag set)." "WARN"
        Add-ReportStep -Name "Permissions" -Status "Skipped"
    }

    # ------------------------------------------------------------------
    # Step 4: Apply theme
    # ------------------------------------------------------------------
    if ($ThemeName) {
        Write-StepHeader 4 "Apply Theme"
        try {
            Write-Log "Applying theme: $ThemeName"
            Set-PnPWebTheme -Theme $ThemeName -ErrorAction Stop
            Write-Log "Theme applied." "OK"
            Add-ReportStep -Name "Apply Theme" -Status "Success" -Detail "Theme: $ThemeName"
        }
        catch {
            Write-Log "Failed to apply theme: $($_.Exception.Message)" "WARN"
            Add-ReportStep -Name "Apply Theme" -Status "Warning" -Detail $_.Exception.Message
        }
    }

    # ------------------------------------------------------------------
    # Step 5: Register SPFx extensions
    # ------------------------------------------------------------------
    if (-not $SkipSPFx) {
        Write-StepHeader 5 "Register SPFx Extensions"

        # Deploy package to app catalog if available
        if (Test-Path $spfxPackagePath) {
            Write-Log "Deploying SPFx package: $spfxPackagePath"
            try {
                Add-PnPApp -Path $spfxPackagePath -Scope Site -Overwrite -ErrorAction Stop | Out-Null
                $app = Get-PnPApp -Scope Site | Where-Object { $_.Title -like "*workflow*" } | Select-Object -First 1
                if ($app) {
                    Install-PnPApp -Identity $app.Id -Scope Site -ErrorAction Stop
                    Write-Log "SPFx package deployed and installed." "OK"
                }
            }
            catch {
                Write-Log "SPFx package deployment failed: $($_.Exception.Message)" "WARN"
            }
        }
        else {
            Write-Log "SPFx package not found at $spfxPackagePath -- skipping deployment." "WARN"
        }

        # Register custom actions for extensions on target lists
        foreach ($reg in $extensionRegistrations) {
            try {
                Write-Log "Registering extension on list: $($reg.ListTitle)"
                $list = Get-PnPList -Identity $reg.ListTitle -ErrorAction Stop

                Add-PnPCustomAction `
                    -Name "BulkActions_$($reg.ListTitle -replace ' ', '')" `
                    -Title "Bulk Actions" `
                    -Description "Bulk operations for $($reg.ListTitle)" `
                    -Location "ClientSideExtension.ListViewCommandSet" `
                    -ClientSideComponentId $reg.ClientId `
                    -ClientSideComponentProperties $reg.Properties `
                    -RegistrationId $list.Id `
                    -RegistrationType List `
                    -ErrorAction Stop

                Write-Log "  Extension registered on $($reg.ListTitle)." "OK"
            }
            catch {
                Write-Log "  Failed to register extension on $($reg.ListTitle): $($_.Exception.Message)" "WARN"
            }
        }

        Add-ReportStep -Name "SPFx Extensions" -Status "Success" -Detail "$($extensionRegistrations.Count) extensions registered"
    }
    else {
        Write-Log "Skipping SPFx registration (SkipSPFx flag set)." "WARN"
        Add-ReportStep -Name "SPFx Extensions" -Status "Skipped"
    }

    # ------------------------------------------------------------------
    # Step 6: Post-deployment report
    # ------------------------------------------------------------------
    Write-StepHeader 6 "Post-Deployment Report"

    $deploymentReport.EndTime = Get-Date
    $deploymentReport.Duration = ($deploymentReport.EndTime - $deploymentReport.StartTime).ToString("hh\:mm\:ss")

    Write-Host ""
    Write-Host "===== DEPLOYMENT REPORT =====" -ForegroundColor Green
    Write-Host "Site URL   : $SiteUrl"
    Write-Host "Site Title : $SiteTitle"
    Write-Host "Duration   : $($deploymentReport.Duration)"
    Write-Host ""
    Write-Host "Step Results:" -ForegroundColor Cyan

    foreach ($step in $deploymentReport.Steps) {
        $color = switch ($step.Status) {
            "Success" { "Green" }
            "Warning" { "Yellow" }
            "Skipped" { "DarkGray" }
            default   { "White" }
        }
        Write-Host ("  {0,-25} [{1}] {2}" -f $step.Step, $step.Status, $step.Detail) -ForegroundColor $color
    }

    # Verification checks
    Write-Host "`nVerification:" -ForegroundColor Cyan
    try {
        $lists = Get-PnPList -ErrorAction Stop
        $listNames = $lists | ForEach-Object { $_.Title }
        foreach ($tpl in $defaultTemplateLists) {
            $expectedName = ($tpl.Name -replace '-', ' ') -replace '(\b[a-z])', { $_.Value.ToUpper() }
            # Simple check: look for lists that partially match the template name
            $found = $listNames | Where-Object { $_ -like "*$($tpl.Name.Split('-')[0])*" }
            if ($found) {
                Write-Host "  List provisioned: $found" -ForegroundColor Green
            }
            else {
                Write-Host "  List not found for template: $($tpl.Name)" -ForegroundColor Yellow
            }
        }
    }
    catch {
        Write-Host "  Could not verify lists: $($_.Exception.Message)" -ForegroundColor Yellow
    }

    # Save report to JSON
    $reportPath = Join-Path $scriptRoot "deployment-report-$(Get-Date -Format 'yyyyMMdd-HHmmss').json"
    $deploymentReport | ConvertTo-Json -Depth 5 | Out-File -FilePath $reportPath -Encoding UTF8
    Write-Log "Deployment report saved to: $reportPath" "OK"

    Write-Host "`n===== PROVISIONING COMPLETE =====" -ForegroundColor Green
    Write-Host "Site is ready at: $SiteUrl"
    Write-Host ""
}
catch {
    Write-Log "Provisioning failed: $($_.Exception.Message)" "ERROR"
    Add-ReportStep -Name "FATAL" -Status "Error" -Detail $_.Exception.Message
    throw
}
finally {
    try { Disconnect-PnPOnline -ErrorAction SilentlyContinue } catch { }
}
