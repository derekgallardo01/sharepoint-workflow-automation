<#
.SYNOPSIS
    Monitors Power Automate flow health and generates an HTML dashboard report.

.DESCRIPTION
    Checks all Power Automate cloud flow run statuses within a specified time window,
    reports failed runs with error details, calculates per-flow success rates, identifies
    stale flows (no recent runs), and generates a self-contained HTML health dashboard.

    Requires the Microsoft.PowerApps.Administration.PowerShell module or
    Power Platform CLI (pac) with appropriate permissions.

.PARAMETER EnvironmentId
    The Power Platform environment ID containing the flows to monitor.
    Use Get-AdminPowerAppEnvironment to find your environment ID.

.PARAMETER Days
    Number of days to look back for flow run history. Default: 7.

.PARAMETER OutputPath
    Path to save the HTML dashboard report. Default: ./FlowHealthReport.html

.PARAMETER FlowFilter
    Optional wildcard filter for flow display names. Default: * (all flows).

.PARAMETER IncludeSuccessDetails
    If specified, includes individual success run details (verbose). Default: $false.

.PARAMETER SendEmail
    If specified, sends the report via email using an SMTP relay.

.PARAMETER SmtpServer
    SMTP server hostname (required if -SendEmail is used).

.PARAMETER EmailTo
    Recipient email addresses (required if -SendEmail is used).

.PARAMETER EmailFrom
    Sender email address (required if -SendEmail is used).

.EXAMPLE
    .\Monitor-FlowHealth.ps1 -EnvironmentId "a1b2c3d4-e5f6-7890-abcd-ef1234567890" -Days 7

.EXAMPLE
    .\Monitor-FlowHealth.ps1 -EnvironmentId $envId -Days 30 -OutputPath "C:\Reports\flow-health.html" -FlowFilter "Workflow*"

.EXAMPLE
    .\Monitor-FlowHealth.ps1 -EnvironmentId $envId -SendEmail -SmtpServer "smtp.contoso.com" -EmailTo "ops@contoso.com" -EmailFrom "flowmonitor@contoso.com"

.NOTES
    Author:  SharePoint Workflow Automation Project
    Version: 1.0.0
    Requires: Microsoft.PowerApps.Administration.PowerShell or Power Automate Management connector
#>

[CmdletBinding()]
param(
    [Parameter(Mandatory = $true, HelpMessage = "Power Platform environment ID")]
    [string]$EnvironmentId,

    [Parameter(Mandatory = $false)]
    [ValidateRange(1, 365)]
    [int]$Days = 7,

    [Parameter(Mandatory = $false)]
    [string]$OutputPath = ".\FlowHealthReport.html",

    [Parameter(Mandatory = $false)]
    [string]$FlowFilter = "*",

    [Parameter(Mandatory = $false)]
    [switch]$IncludeSuccessDetails,

    [Parameter(Mandatory = $false)]
    [switch]$SendEmail,

    [Parameter(Mandatory = $false)]
    [string]$SmtpServer,

    [Parameter(Mandatory = $false)]
    [string[]]$EmailTo,

    [Parameter(Mandatory = $false)]
    [string]$EmailFrom
)

$ErrorActionPreference = "Stop"
Set-StrictMode -Version Latest

# ---------------------------------------------------------------------------
# Constants
# ---------------------------------------------------------------------------

$Script:HealthThresholds = @{
    SuccessRateGreen  = 95   # >= 95% = Healthy
    SuccessRateYellow = 80   # >= 80% = Warning
                              # < 80%  = Critical
    StaleFlowDays     = 3    # No runs in N days = Stale
}

# ---------------------------------------------------------------------------
# Helper Functions
# ---------------------------------------------------------------------------

function Write-Status {
    param([string]$Message, [string]$Level = "Info")
    $prefix = switch ($Level) {
        "Info"    { "[INFO]" }
        "Warning" { "[WARN]" }
        "Error"   { "[ERROR]" }
        "Success" { "[OK]" }
    }
    $color = switch ($Level) {
        "Info"    { "Cyan" }
        "Warning" { "Yellow" }
        "Error"   { "Red" }
        "Success" { "Green" }
    }
    Write-Host "$prefix $Message" -ForegroundColor $color
}

function Get-HealthStatus {
    param([double]$SuccessRate)
    if ($SuccessRate -ge $Script:HealthThresholds.SuccessRateGreen) { return "Healthy" }
    elseif ($SuccessRate -ge $Script:HealthThresholds.SuccessRateYellow) { return "Warning" }
    else { return "Critical" }
}

function Get-HealthColor {
    param([string]$Status)
    switch ($Status) {
        "Healthy"  { return "#10b981" }
        "Warning"  { return "#f59e0b" }
        "Critical" { return "#ef4444" }
        "Stale"    { return "#6b7280" }
        default    { return "#6b7280" }
    }
}

function Get-HealthEmoji {
    param([string]$Status)
    switch ($Status) {
        "Healthy"  { return "&#x2705;" }  # Green check
        "Warning"  { return "&#x26A0;" }  # Warning triangle
        "Critical" { return "&#x274C;" }  # Red X
        "Stale"    { return "&#x23F8;" }  # Pause
        default    { return "&#x2753;" }  # Question mark
    }
}

# ---------------------------------------------------------------------------
# Module Validation
# ---------------------------------------------------------------------------

Write-Status "Validating prerequisites..."

$moduleAvailable = Get-Module -ListAvailable -Name "Microsoft.PowerApps.Administration.PowerShell"
if (-not $moduleAvailable) {
    Write-Status "Microsoft.PowerApps.Administration.PowerShell module not found. Attempting install..." "Warning"
    try {
        Install-Module -Name Microsoft.PowerApps.Administration.PowerShell -Scope CurrentUser -Force -AllowClobber
        Write-Status "Module installed successfully." "Success"
    }
    catch {
        Write-Status "Failed to install module: $_" "Error"
        Write-Status "Install manually: Install-Module -Name Microsoft.PowerApps.Administration.PowerShell" "Error"
        exit 1
    }
}

Import-Module Microsoft.PowerApps.Administration.PowerShell -ErrorAction Stop
Write-Status "Module loaded." "Success"

# Validate email parameters if -SendEmail specified
if ($SendEmail) {
    if (-not $SmtpServer -or -not $EmailTo -or -not $EmailFrom) {
        Write-Status "-SendEmail requires -SmtpServer, -EmailTo, and -EmailFrom parameters." "Error"
        exit 1
    }
}

# ---------------------------------------------------------------------------
# Data Collection
# ---------------------------------------------------------------------------

$reportStartDate = (Get-Date).AddDays(-$Days)
$reportEndDate = Get-Date
$reportTimestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

Write-Status "Collecting flow data for environment: $EnvironmentId"
Write-Status "Time window: $($reportStartDate.ToString('yyyy-MM-dd')) to $($reportEndDate.ToString('yyyy-MM-dd')) ($Days days)"

# Get all flows in the environment
try {
    $allFlows = Get-AdminFlow -EnvironmentName $EnvironmentId -ErrorAction Stop |
        Where-Object { $_.DisplayName -like $FlowFilter }
    Write-Status "Found $($allFlows.Count) flows matching filter '$FlowFilter'." "Success"
}
catch {
    Write-Status "Failed to retrieve flows: $_" "Error"
    exit 1
}

if ($allFlows.Count -eq 0) {
    Write-Status "No flows found matching filter '$FlowFilter'. Exiting." "Warning"
    exit 0
}

# ---------------------------------------------------------------------------
# Flow Analysis
# ---------------------------------------------------------------------------

$flowReports = [System.Collections.ArrayList]::new()

foreach ($flow in $allFlows) {
    $flowName = $flow.DisplayName
    $flowId = $flow.FlowName
    Write-Status "Analyzing: $flowName ($flowId)"

    $flowReport = [PSCustomObject]@{
        FlowName       = $flowName
        FlowId         = $flowId
        State          = $flow.Enabled ? "Enabled" : "Disabled"
        TotalRuns      = 0
        Succeeded      = 0
        Failed         = 0
        Cancelled       = 0
        TimedOut       = 0
        SuccessRate    = 0.0
        AvgDuration    = "N/A"
        LastRun        = $null
        HealthStatus   = "Unknown"
        IsStale        = $false
        FailedRuns     = [System.Collections.ArrayList]::new()
        RecentRuns     = [System.Collections.ArrayList]::new()
    }

    # Get flow runs within the time window
    try {
        $runs = Get-AdminFlowRun -EnvironmentName $EnvironmentId -FlowName $flowId -ErrorAction Stop |
            Where-Object {
                $_.StartTime -ge $reportStartDate -and $_.StartTime -le $reportEndDate
            }
    }
    catch {
        Write-Status "  Could not retrieve runs for $flowName : $_" "Warning"
        $flowReport.HealthStatus = "Unknown"
        [void]$flowReports.Add($flowReport)
        continue
    }

    $flowReport.TotalRuns = $runs.Count

    if ($runs.Count -eq 0) {
        $flowReport.IsStale = $true
        $flowReport.HealthStatus = "Stale"
        Write-Status "  No runs in the last $Days days (Stale)" "Warning"
        [void]$flowReports.Add($flowReport)
        continue
    }

    # Categorize runs
    $durations = [System.Collections.ArrayList]::new()

    foreach ($run in $runs) {
        $runDetail = [PSCustomObject]@{
            RunId     = $run.FlowRunName
            Status    = $run.Status
            StartTime = $run.StartTime
            EndTime   = $run.EndTime
            Duration  = if ($run.EndTime -and $run.StartTime) {
                ($run.EndTime - $run.StartTime).TotalSeconds
            } else { 0 }
            Error     = $run.Error
        }

        switch ($run.Status) {
            "Succeeded" {
                $flowReport.Succeeded++
                if ($runDetail.Duration -gt 0) {
                    [void]$durations.Add($runDetail.Duration)
                }
            }
            "Failed" {
                $flowReport.Failed++
                [void]$flowReport.FailedRuns.Add($runDetail)
            }
            "Cancelled" {
                $flowReport.Cancelled++
            }
            "TimedOut" {
                $flowReport.TimedOut++
                [void]$flowReport.FailedRuns.Add($runDetail)
            }
        }

        if ($IncludeSuccessDetails -or $run.Status -ne "Succeeded") {
            [void]$flowReport.RecentRuns.Add($runDetail)
        }
    }

    # Calculate metrics
    if ($flowReport.TotalRuns -gt 0) {
        $flowReport.SuccessRate = [math]::Round(
            ($flowReport.Succeeded / $flowReport.TotalRuns) * 100, 1
        )
    }

    if ($durations.Count -gt 0) {
        $avgSeconds = ($durations | Measure-Object -Average).Average
        $ts = [TimeSpan]::FromSeconds($avgSeconds)
        $flowReport.AvgDuration = if ($ts.TotalMinutes -ge 1) {
            "{0:N1} min" -f $ts.TotalMinutes
        } else {
            "{0:N0} sec" -f $ts.TotalSeconds
        }
    }

    $flowReport.LastRun = ($runs | Sort-Object StartTime -Descending | Select-Object -First 1).StartTime
    $flowReport.HealthStatus = Get-HealthStatus -SuccessRate $flowReport.SuccessRate

    # Check if stale (last run was more than StaleFlowDays ago even though runs exist in window)
    if ($flowReport.LastRun -lt (Get-Date).AddDays(-$Script:HealthThresholds.StaleFlowDays)) {
        $flowReport.IsStale = $true
    }

    $statusMsg = "$($flowReport.HealthStatus): $($flowReport.SuccessRate)% success ($($flowReport.Succeeded)/$($flowReport.TotalRuns))"
    $statusLevel = switch ($flowReport.HealthStatus) {
        "Healthy"  { "Success" }
        "Warning"  { "Warning" }
        "Critical" { "Error" }
        default    { "Info" }
    }
    Write-Status "  $statusMsg" $statusLevel

    [void]$flowReports.Add($flowReport)
}

# ---------------------------------------------------------------------------
# Summary Metrics
# ---------------------------------------------------------------------------

$totalFlows     = $flowReports.Count
$healthyFlows   = ($flowReports | Where-Object { $_.HealthStatus -eq "Healthy" }).Count
$warningFlows   = ($flowReports | Where-Object { $_.HealthStatus -eq "Warning" }).Count
$criticalFlows  = ($flowReports | Where-Object { $_.HealthStatus -eq "Critical" }).Count
$staleFlows     = ($flowReports | Where-Object { $_.IsStale }).Count
$unknownFlows   = ($flowReports | Where-Object { $_.HealthStatus -eq "Unknown" }).Count

$totalRuns      = ($flowReports | Measure-Object -Property TotalRuns -Sum).Sum
$totalSucceeded = ($flowReports | Measure-Object -Property Succeeded -Sum).Sum
$totalFailed    = ($flowReports | Measure-Object -Property Failed -Sum).Sum
$overallRate    = if ($totalRuns -gt 0) { [math]::Round(($totalSucceeded / $totalRuns) * 100, 1) } else { 0 }
$overallHealth  = Get-HealthStatus -SuccessRate $overallRate

Write-Status ""
Write-Status "===== SUMMARY ====="
Write-Status "Total Flows: $totalFlows | Healthy: $healthyFlows | Warning: $warningFlows | Critical: $criticalFlows | Stale: $staleFlows"
Write-Status "Total Runs: $totalRuns | Succeeded: $totalSucceeded | Failed: $totalFailed | Overall: $overallRate%"

# ---------------------------------------------------------------------------
# HTML Dashboard Generation
# ---------------------------------------------------------------------------

Write-Status ""
Write-Status "Generating HTML dashboard..."

$failedRunsHtml = ""
foreach ($flow in ($flowReports | Where-Object { $_.FailedRuns.Count -gt 0 })) {
    foreach ($run in $flow.FailedRuns) {
        $errorText = if ($run.Error) {
            [System.Web.HttpUtility]::HtmlEncode($run.Error.ToString().Substring(0, [Math]::Min(200, $run.Error.ToString().Length)))
        } else {
            "No error details available"
        }
        $failedRunsHtml += @"
                <tr>
                    <td>$([System.Web.HttpUtility]::HtmlEncode($flow.FlowName))</td>
                    <td><code>$($run.RunId)</code></td>
                    <td>$($run.StartTime.ToString('yyyy-MM-dd HH:mm'))</td>
                    <td><span class="badge badge-$($run.Status.ToLower())">$($run.Status)</span></td>
                    <td class="error-cell">$errorText</td>
                </tr>
"@
    }
}

$flowRowsHtml = ""
foreach ($flow in ($flowReports | Sort-Object SuccessRate)) {
    $healthColor = Get-HealthColor -Status $flow.HealthStatus
    $healthEmoji = Get-HealthEmoji -Status $flow.HealthStatus
    $staleTag = if ($flow.IsStale) { ' <span class="badge badge-stale">STALE</span>' } else { "" }
    $lastRunText = if ($flow.LastRun) { $flow.LastRun.ToString('yyyy-MM-dd HH:mm') } else { "Never" }

    $rateBarWidth = [Math]::Max($flow.SuccessRate, 2) # Minimum width for visibility
    $rateBarColor = $healthColor

    $flowRowsHtml += @"
                <tr>
                    <td>
                        <span style="color: $healthColor;">$healthEmoji</span>
                        $([System.Web.HttpUtility]::HtmlEncode($flow.FlowName))$staleTag
                    </td>
                    <td><span class="badge badge-$($flow.State.ToLower())">$($flow.State)</span></td>
                    <td>$($flow.TotalRuns)</td>
                    <td style="color: #10b981;">$($flow.Succeeded)</td>
                    <td style="color: #ef4444;">$($flow.Failed)</td>
                    <td>
                        <div class="rate-bar-container">
                            <div class="rate-bar" style="width: $($rateBarWidth)%; background: $rateBarColor;"></div>
                        </div>
                        <span style="color: $healthColor; font-weight: 600;">$($flow.SuccessRate)%</span>
                    </td>
                    <td>$($flow.AvgDuration)</td>
                    <td>$lastRunText</td>
                </tr>
"@
}

$htmlReport = @"
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Power Automate Flow Health Dashboard</title>
    <style>
        * { margin: 0; padding: 0; box-sizing: border-box; }
        body {
            font-family: 'Segoe UI', -apple-system, sans-serif;
            background: #0f172a;
            color: #e2e8f0;
            padding: 24px;
            line-height: 1.5;
        }
        .container { max-width: 1400px; margin: 0 auto; }
        h1 { font-size: 24px; font-weight: 700; margin-bottom: 4px; }
        h2 { font-size: 18px; font-weight: 600; margin: 32px 0 16px; color: #f1f5f9; }
        .subtitle { font-size: 14px; color: #64748b; margin-bottom: 24px; }
        .card {
            background: rgba(30, 41, 59, 0.8);
            border: 1px solid rgba(148, 163, 184, 0.1);
            border-radius: 12px;
            padding: 20px;
            margin-bottom: 16px;
        }
        .summary-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(160px, 1fr));
            gap: 12px;
            margin-bottom: 24px;
        }
        .metric-card {
            background: rgba(15, 23, 42, 0.6);
            border: 1px solid rgba(148, 163, 184, 0.08);
            border-radius: 10px;
            padding: 16px;
            text-align: center;
        }
        .metric-value { font-size: 32px; font-weight: 700; line-height: 1.2; }
        .metric-label { font-size: 12px; color: #64748b; margin-top: 4px; text-transform: uppercase; letter-spacing: 0.05em; }
        table { width: 100%; border-collapse: collapse; }
        th { text-align: left; padding: 10px 14px; font-size: 11px; font-weight: 600; color: #64748b; text-transform: uppercase; letter-spacing: 0.05em; border-bottom: 1px solid rgba(148,163,184,0.1); }
        td { padding: 10px 14px; font-size: 13px; border-bottom: 1px solid rgba(148,163,184,0.04); }
        tr:hover { background: rgba(148,163,184,0.03); }
        .badge {
            display: inline-block;
            padding: 2px 8px;
            border-radius: 9999px;
            font-size: 11px;
            font-weight: 600;
        }
        .badge-healthy, .badge-succeeded, .badge-enabled { background: rgba(16,185,129,0.12); color: #10b981; }
        .badge-warning { background: rgba(245,158,11,0.12); color: #f59e0b; }
        .badge-critical, .badge-failed { background: rgba(239,68,68,0.12); color: #ef4444; }
        .badge-stale, .badge-disabled { background: rgba(107,114,128,0.12); color: #6b7280; }
        .badge-timedout { background: rgba(245,158,11,0.12); color: #f59e0b; }
        .badge-cancelled { background: rgba(107,114,128,0.12); color: #9ca3af; }
        .rate-bar-container { width: 80px; height: 6px; background: rgba(148,163,184,0.1); border-radius: 3px; display: inline-block; vertical-align: middle; margin-right: 8px; }
        .rate-bar { height: 100%; border-radius: 3px; transition: width 0.3s; }
        .error-cell { max-width: 300px; font-size: 12px; color: #94a3b8; word-break: break-word; }
        code { font-family: 'Cascadia Code', 'Consolas', monospace; font-size: 11px; background: rgba(148,163,184,0.08); padding: 1px 5px; border-radius: 4px; }
        .footer { margin-top: 32px; text-align: center; font-size: 12px; color: #475569; }
    </style>
</head>
<body>
    <div class="container">
        <h1>Power Automate Flow Health Dashboard</h1>
        <div class="subtitle">
            Environment: <code>$EnvironmentId</code> &bull;
            Period: $($reportStartDate.ToString('MMM dd, yyyy')) &ndash; $($reportEndDate.ToString('MMM dd, yyyy')) ($Days days) &bull;
            Generated: $reportTimestamp
        </div>

        <!-- Summary Metrics -->
        <div class="summary-grid">
            <div class="metric-card">
                <div class="metric-value" style="color: $(Get-HealthColor -Status $overallHealth);">$overallRate%</div>
                <div class="metric-label">Overall Success Rate</div>
            </div>
            <div class="metric-card">
                <div class="metric-value" style="color: #f1f5f9;">$totalFlows</div>
                <div class="metric-label">Total Flows</div>
            </div>
            <div class="metric-card">
                <div class="metric-value" style="color: #f1f5f9;">$totalRuns</div>
                <div class="metric-label">Total Runs</div>
            </div>
            <div class="metric-card">
                <div class="metric-value" style="color: #10b981;">$healthyFlows</div>
                <div class="metric-label">Healthy Flows</div>
            </div>
            <div class="metric-card">
                <div class="metric-value" style="color: #f59e0b;">$warningFlows</div>
                <div class="metric-label">Warning Flows</div>
            </div>
            <div class="metric-card">
                <div class="metric-value" style="color: #ef4444;">$criticalFlows</div>
                <div class="metric-label">Critical Flows</div>
            </div>
            <div class="metric-card">
                <div class="metric-value" style="color: #6b7280;">$staleFlows</div>
                <div class="metric-label">Stale Flows</div>
            </div>
            <div class="metric-card">
                <div class="metric-value" style="color: #ef4444;">$totalFailed</div>
                <div class="metric-label">Failed Runs</div>
            </div>
        </div>

        <!-- Per-Flow Breakdown -->
        <h2>Flow Status Breakdown</h2>
        <div class="card">
            <table>
                <thead>
                    <tr>
                        <th>Flow Name</th>
                        <th>State</th>
                        <th>Runs</th>
                        <th>Passed</th>
                        <th>Failed</th>
                        <th>Success Rate</th>
                        <th>Avg Duration</th>
                        <th>Last Run</th>
                    </tr>
                </thead>
                <tbody>
$flowRowsHtml
                </tbody>
            </table>
        </div>

        <!-- Failed Runs Detail -->
        <h2>Failed Run Details</h2>
        <div class="card">
            $(if ($failedRunsHtml) {
                @"
            <table>
                <thead>
                    <tr>
                        <th>Flow</th>
                        <th>Run ID</th>
                        <th>Time</th>
                        <th>Status</th>
                        <th>Error</th>
                    </tr>
                </thead>
                <tbody>
$failedRunsHtml
                </tbody>
            </table>
"@
            } else {
                '<p style="color: #10b981; padding: 12px;">&#x2705; No failed runs in the reporting period.</p>'
            })
        </div>

        <div class="footer">
            Generated by Monitor-FlowHealth.ps1 &bull; SharePoint Workflow Automation Project
        </div>
    </div>
</body>
</html>
"@

# ---------------------------------------------------------------------------
# Output
# ---------------------------------------------------------------------------

# Ensure output directory exists
$outputDir = Split-Path -Path $OutputPath -Parent
if ($outputDir -and -not (Test-Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
}

$htmlReport | Out-File -FilePath $OutputPath -Encoding UTF8 -Force
Write-Status "Dashboard saved to: $OutputPath" "Success"

# ---------------------------------------------------------------------------
# Email (Optional)
# ---------------------------------------------------------------------------

if ($SendEmail) {
    Write-Status "Sending report via email..."
    try {
        $subject = "Power Automate Health Report - $overallHealth ($overallRate%) - $(Get-Date -Format 'yyyy-MM-dd')"
        Send-MailMessage `
            -SmtpServer $SmtpServer `
            -From $EmailFrom `
            -To $EmailTo `
            -Subject $subject `
            -Body $htmlReport `
            -BodyAsHtml `
            -Priority $(if ($criticalFlows -gt 0) { "High" } else { "Normal" }) `
            -ErrorAction Stop
        Write-Status "Email sent to: $($EmailTo -join ', ')" "Success"
    }
    catch {
        Write-Status "Failed to send email: $_" "Error"
    }
}

# ---------------------------------------------------------------------------
# Return object for pipeline use
# ---------------------------------------------------------------------------

Write-Status ""
Write-Status "Monitor-FlowHealth complete." "Success"

return [PSCustomObject]@{
    Timestamp     = $reportTimestamp
    EnvironmentId = $EnvironmentId
    Days          = $Days
    OverallHealth = $overallHealth
    OverallRate   = $overallRate
    TotalFlows    = $totalFlows
    TotalRuns     = $totalRuns
    Healthy       = $healthyFlows
    Warning       = $warningFlows
    Critical      = $criticalFlows
    Stale         = $staleFlows
    TotalFailed   = $totalFailed
    ReportPath    = (Resolve-Path $OutputPath).Path
    FlowReports   = $flowReports
}
