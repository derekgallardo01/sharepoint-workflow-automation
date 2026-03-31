#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    Pester v5 tests for the provisioning PowerShell scripts.

.DESCRIPTION
    Validates that each provisioning script exists, has correct CmdletBinding,
    expected parameters, proper help comments, SupportsShouldProcess where
    applicable, and error handling.

.EXAMPLE
    Invoke-Pester -Path .\tests\Test-Provisioning.ps1
#>

BeforeAll {
    $scriptRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $provisioningDir = Join-Path $scriptRoot "provisioning"

    if (-not (Test-Path $provisioningDir)) {
        $provisioningDir = Join-Path (Split-Path -Parent $PSScriptRoot) "provisioning"
    }
}

Describe "Provisioning Scripts" {

    Context "Deploy-WorkflowSolution.ps1" {
        BeforeAll {
            $scriptPath = Join-Path $provisioningDir "Deploy-WorkflowSolution.ps1"
        }

        It "Should exist at the expected path" {
            $scriptPath | Should -Exist
        }

        It "Should have CmdletBinding attribute" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[CmdletBinding'
        }

        It "Should support -WhatIf (SupportsShouldProcess)" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'SupportsShouldProcess'
        }

        It "Should have mandatory parameter SiteUrl" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'Mandatory\s*=\s*\$true'
            $content | Should -Match '\$SiteUrl'
        }

        It "Should validate SiteUrl is a SharePoint URL" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'ValidatePattern.*sharepoint\.com'
        }

        It "Should have optional parameter TenantAdminUrl" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\$TenantAdminUrl'
        }

        It "Should have optional parameter Credential" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[PSCredential\]\$Credential'
        }

        It "Should have switch parameter CreateSite" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[switch\]\$CreateSite'
        }

        It "Should have switch parameters SkipListTemplates and SkipSPFx" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[switch\]\$SkipListTemplates'
            $content | Should -Match '\[switch\]\$SkipSPFx'
        }

        It "Should have .SYNOPSIS help comment" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.SYNOPSIS'
        }

        It "Should have .DESCRIPTION help comment" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.DESCRIPTION'
        }

        It "Should have .PARAMETER help comments" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.PARAMETER\s+SiteUrl'
            $content | Should -Match '\.PARAMETER\s+TenantAdminUrl'
            $content | Should -Match '\.PARAMETER\s+Credential'
        }

        It "Should have .EXAMPLE help comments" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.EXAMPLE'
        }

        It "Should have error handling (try/catch)" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\btry\b'
            $content | Should -Match '\bcatch\b'
        }

        It "Should set ErrorActionPreference to Stop" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match "\`\$ErrorActionPreference\s*=\s*['""]Stop['""]"
        }

        It "Should reference list template files" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'project-tracker\.xml'
            $content | Should -Match 'document-approval\.xml'
            $content | Should -Match 'change-request\.xml'
            $content | Should -Match 'employee-onboarding\.xml'
        }

        It "Should disconnect from SharePoint in a finally block" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\bfinally\b'
            $content | Should -Match 'Disconnect-PnPOnline'
        }
    }

    Context "Remove-WorkflowSolution.ps1" {
        BeforeAll {
            $scriptPath = Join-Path $provisioningDir "Remove-WorkflowSolution.ps1"
        }

        It "Should exist at the expected path" {
            $scriptPath | Should -Exist
        }

        It "Should have CmdletBinding attribute" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[CmdletBinding'
        }

        It "Should support -WhatIf (SupportsShouldProcess)" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'SupportsShouldProcess'
        }

        It "Should have ConfirmImpact set to High" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match "ConfirmImpact\s*=\s*['""]High['""]"
        }

        It "Should have mandatory parameter SiteUrl" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'Mandatory\s*=\s*\$true'
            $content | Should -Match '\$SiteUrl'
        }

        It "Should validate SiteUrl is a SharePoint URL" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'ValidatePattern.*sharepoint\.com'
        }

        It "Should have switch parameter RemoveLists" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[switch\]\$RemoveLists'
        }

        It "Should have switch parameter RemoveApp" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[switch\]\$RemoveApp'
        }

        It "Should have switch parameter RemoveSite" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[switch\]\$RemoveSite'
        }

        It "Should have switch parameter Force" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[switch\]\$Force'
        }

        It "Should have .SYNOPSIS help comment" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.SYNOPSIS'
        }

        It "Should have .DESCRIPTION help comment with warning about destructive operation" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.DESCRIPTION'
            $content | Should -Match 'WARNING|destructive'
        }

        It "Should have .EXAMPLE help comments" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.EXAMPLE'
        }

        It "Should have error handling (try/catch)" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\btry\b'
            $content | Should -Match '\bcatch\b'
        }

        It "Should reference the four provisioned lists for cleanup" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'Project Tracker'
            $content | Should -Match 'Document Approval Queue'
            $content | Should -Match 'Change Request Log'
            $content | Should -Match 'Employee Onboarding Checklist'
        }

        It "Should disconnect from SharePoint in a finally block" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\bfinally\b'
            $content | Should -Match 'Disconnect-PnPOnline'
        }
    }

    Context "Set-ListPermissions.ps1" {
        BeforeAll {
            $scriptPath = Join-Path $provisioningDir "Set-ListPermissions.ps1"
        }

        It "Should exist at the expected path" {
            $scriptPath | Should -Exist
        }

        It "Should have CmdletBinding attribute" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[CmdletBinding'
        }

        It "Should support -WhatIf (SupportsShouldProcess)" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'SupportsShouldProcess'
        }

        It "Should have mandatory parameter SiteUrl" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'Mandatory\s*=\s*\$true'
            $content | Should -Match '\$SiteUrl'
        }

        It "Should validate SiteUrl is a SharePoint URL" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'ValidatePattern.*sharepoint\.com'
        }

        It "Should support two parameter sets: Single and Bulk" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match "ParameterSetName\s*=\s*['""]Single['""]"
            $content | Should -Match "ParameterSetName\s*=\s*['""]Bulk['""]"
        }

        It "Should have parameter ListName in Single parameter set" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\$ListName'
        }

        It "Should have parameter GroupName in Single parameter set" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\$GroupName'
        }

        It "Should have parameter PermissionLevel with ValidateSet" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'ValidateSet.*Read.*Contribute.*Edit'
            $content | Should -Match '\$PermissionLevel'
        }

        It "Should have parameter ConfigFile for Bulk mode" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\$ConfigFile'
        }

        It "Should have switch parameter BreakInheritance" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\[switch\]\$BreakInheritance'
        }

        It "Should have .SYNOPSIS help comment" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.SYNOPSIS'
        }

        It "Should have .DESCRIPTION help comment" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.DESCRIPTION'
        }

        It "Should have .PARAMETER help comments" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.PARAMETER\s+SiteUrl'
            $content | Should -Match '\.PARAMETER\s+ListName'
            $content | Should -Match '\.PARAMETER\s+GroupName'
            $content | Should -Match '\.PARAMETER\s+PermissionLevel'
        }

        It "Should have .EXAMPLE help comments" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\.EXAMPLE'
        }

        It "Should have error handling (try/catch)" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match '\btry\b'
            $content | Should -Match '\bcatch\b'
        }

        It "Should set StrictMode" {
            $content = Get-Content -Path $scriptPath -Raw
            $content | Should -Match 'Set-StrictMode'
        }
    }
}
