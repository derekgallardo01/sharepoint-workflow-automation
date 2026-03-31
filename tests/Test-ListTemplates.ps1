#Requires -Modules @{ ModuleName = 'Pester'; ModuleVersion = '5.0' }

<#
.SYNOPSIS
    Pester v5 tests for the PnP list template XML files.

.DESCRIPTION
    Validates that each XML template file is well-formed XML and contains
    the required PnP provisioning elements: ListInstance, Fields/FieldRefs,
    ContentTypes, and Views.

.EXAMPLE
    Invoke-Pester -Path .\tests\Test-ListTemplates.ps1
#>

BeforeAll {
    $scriptRoot = Split-Path -Parent (Split-Path -Parent $PSScriptRoot)
    $templatesDir = Join-Path $scriptRoot "list-templates"

    if (-not (Test-Path $templatesDir)) {
        $templatesDir = Join-Path (Split-Path -Parent $PSScriptRoot) "list-templates"
    }

    $templateFiles = @(
        "project-tracker.xml",
        "document-approval.xml",
        "change-request.xml",
        "employee-onboarding.xml",
        "it-asset-tracking.xml"
    )
}

Describe "List Template XML Files" {

    Context "Template files exist" {
        It "Should have the list-templates directory" {
            $templatesDir | Should -Exist
        }

        It "Should contain <templateFile> template file" -ForEach @(
            @{ templateFile = "project-tracker.xml" },
            @{ templateFile = "document-approval.xml" },
            @{ templateFile = "change-request.xml" },
            @{ templateFile = "employee-onboarding.xml" },
            @{ templateFile = "it-asset-tracking.xml" }
        ) {
            $filePath = Join-Path $templatesDir $templateFile
            $filePath | Should -Exist
        }
    }

    Context "XML validity for <templateFile>" -ForEach @(
        @{ templateFile = "project-tracker.xml" },
        @{ templateFile = "document-approval.xml" },
        @{ templateFile = "change-request.xml" },
        @{ templateFile = "employee-onboarding.xml" },
        @{ templateFile = "it-asset-tracking.xml" }
    ) {
        BeforeAll {
            $filePath = Join-Path $templatesDir $templateFile
            $xmlContent = $null
            $parseError = $null
            try {
                [xml]$xmlContent = Get-Content -Path $filePath -Raw
            }
            catch {
                $parseError = $_.Exception.Message
            }
        }

        It "Should be valid XML (parseable without errors)" {
            $parseError | Should -BeNullOrEmpty
            $xmlContent | Should -Not -BeNullOrEmpty
        }

        It "Should have an XML declaration or root element" {
            $xmlContent | Should -Not -BeNullOrEmpty
            $xmlContent.DocumentElement | Should -Not -BeNullOrEmpty
        }
    }

    Context "PnP schema elements in <templateFile>" -ForEach @(
        @{ templateFile = "project-tracker.xml" },
        @{ templateFile = "document-approval.xml" },
        @{ templateFile = "change-request.xml" },
        @{ templateFile = "employee-onboarding.xml" },
        @{ templateFile = "it-asset-tracking.xml" }
    ) {
        BeforeAll {
            $filePath = Join-Path $templatesDir $templateFile
            $rawContent = Get-Content -Path $filePath -Raw
        }

        It "Should contain a ListInstance element" {
            $rawContent | Should -Match 'ListInstance'
        }

        It "Should contain Field definitions (SiteFields or Field elements)" {
            # PnP templates define fields via <pnp:SiteFields> or <Field> elements
            $rawContent | Should -Match '<Field\b|SiteFields'
        }

        It "Should contain FieldRef elements (in ContentTypes or Views)" {
            $rawContent | Should -Match 'FieldRef'
        }

        It "Should contain at least one View element" {
            $rawContent | Should -Match '<View\b'
        }

        It "Should reference the PnP provisioning namespace" {
            $rawContent | Should -Match 'schemas\.dev\.office\.com/PnP|ProvisioningSchema'
        }

        It "Should have a ProvisioningTemplate element with an ID" {
            $rawContent | Should -Match 'ProvisioningTemplate\s+ID='
        }
    }

    Context "project-tracker.xml specific content" {
        BeforeAll {
            $filePath = Join-Path $templatesDir "project-tracker.xml"
            $rawContent = Get-Content -Path $filePath -Raw
        }

        It "Should define ProjectStatus field as a Choice type" {
            $rawContent | Should -Match 'Name="ProjectStatus"'
            $rawContent | Should -Match 'Type="Choice"'
        }

        It "Should define ProjectPriority field" {
            $rawContent | Should -Match 'Name="ProjectPriority"'
        }

        It "Should define ProjectAssignedTo field as a User type" {
            $rawContent | Should -Match 'Name="ProjectAssignedTo"'
            $rawContent | Should -Match 'Type="User"'
        }

        It "Should have ContentType bindings for the list" {
            $rawContent | Should -Match 'ContentTypeBinding'
        }

        It "Should define multiple views (All Items, My Projects, Active, Overdue, By Status)" {
            $rawContent | Should -Match 'All Items'
            $rawContent | Should -Match 'My Projects'
            $rawContent | Should -Match 'Active Projects'
            $rawContent | Should -Match 'Overdue'
            $rawContent | Should -Match 'By Status'
        }
    }
}
