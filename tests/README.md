# Testing

## Test Categories

### Provisioning Script Tests (`Test-Provisioning.ps1`)

Pester v5 tests that validate the provisioning PowerShell scripts without requiring a live SharePoint environment. Tests verify:

- Scripts exist at expected paths
- CmdletBinding attributes and parameter declarations are correct
- SupportsShouldProcess (WhatIf) is enabled where applicable
- Parameter validation (ValidatePattern, ValidateSet) is present
- Help comments (.SYNOPSIS, .DESCRIPTION, .PARAMETER, .EXAMPLE) are present
- Error handling patterns (try/catch, ErrorActionPreference, finally) are in place

### List Template Tests (`Test-ListTemplates.ps1`)

Validates the PnP provisioning XML template files:

- Each XML file is well-formed (parseable without errors)
- Each has a ListInstance element
- Each has Field definitions (SiteFields)
- Each has FieldRef elements
- Each has View definitions
- Each references the PnP provisioning namespace
- Specific field definitions and view configurations for project-tracker.xml

## Prerequisites

- **PowerShell 5.1+** or **PowerShell 7+**
- **Pester v5+**: Install with `Install-Module Pester -MinimumVersion 5.0 -Scope CurrentUser`

## Running Tests

```powershell
# Run all tests
Invoke-Pester -Path .\tests\

# Run with detailed output
Invoke-Pester -Path .\tests\ -Output Detailed

# Run only provisioning tests
Invoke-Pester -Path .\tests\Test-Provisioning.ps1

# Run only XML template tests
Invoke-Pester -Path .\tests\Test-ListTemplates.ps1

# Generate NUnit XML report for CI
Invoke-Pester -Path .\tests\ -OutputFormat NUnitXml -OutputFile .\tests\results.xml
```

## Test Categories Overview

| Category | File | Requires Live Environment |
|---|---|---|
| Script validation | `Test-Provisioning.ps1` | No |
| XML template validation | `Test-ListTemplates.ps1` | No |
| Integration testing | (manual) | Yes (SharePoint Online tenant) |

## CI/CD Integration

### Azure DevOps Pipeline example

```yaml
steps:
  - task: PowerShell@2
    displayName: 'Run Pester Tests'
    inputs:
      targetType: 'inline'
      script: |
        Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser
        $results = Invoke-Pester -Path .\tests\ -OutputFormat NUnitXml -OutputFile .\tests\results.xml -PassThru
        if ($results.FailedCount -gt 0) { exit 1 }
      pwsh: true

  - task: PublishTestResults@2
    inputs:
      testResultsFormat: 'NUnit'
      testResultsFiles: 'tests/results.xml'
    condition: always()
```

### GitHub Actions example

```yaml
- name: Run Pester Tests
  shell: pwsh
  run: |
    Install-Module Pester -MinimumVersion 5.0 -Force -Scope CurrentUser
    Invoke-Pester -Path .\tests\ -Output Detailed -CI
```

## Notes

- All tests run offline and do not require a SharePoint connection.
- Provisioning scripts support `-WhatIf` for dry-run testing in live environments.
- The PnP.PowerShell module is only needed for actual deployments, not for running these tests.
