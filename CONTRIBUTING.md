# Contributing

Thank you for your interest in contributing to the SharePoint Workflow Automation project.

## Prerequisites

- **Node.js** 18.x LTS
- **npm** 9+
- **Gulp CLI**: `npm install -g gulp-cli`
- **PnP.PowerShell** 2.x+: `Install-Module PnP.PowerShell`
- A **SharePoint Online** tenant with an App Catalog (for testing SPFx and list template changes)
- A **Power Automate** licence with premium connectors (for testing flow changes)

## Setup

```bash
# Clone and install
git clone https://github.com/your-org/sharepoint-workflow-automation.git
cd sharepoint-workflow-automation/spfx-extensions
npm install

# Start the local workbench
gulp serve
```

## Development Workflow

1. Create a feature branch from `main`: `git checkout -b feature/your-change`
2. Make your changes.
3. Test locally:
   - **SPFx extensions**: `gulp serve` and test against a dev SharePoint site using the debug query string.
   - **List templates**: Deploy to a dev site with `Invoke-PnPSiteTemplate -Path your-template.xml`.
   - **PowerShell scripts**: Run with `-WhatIf` first, then test against a dev site.
   - **Power Automate flows**: Validate JSON, then test-import in a dev environment.
4. Verify the build passes: `gulp build`
5. Commit with a clear message and open a Pull Request against `main`.

## Code Style

- **TypeScript**: Follow existing patterns. Use Fluent UI React components. Use PnP SPFx controls where applicable.
- **SPFx patterns**: ListView Command Sets should use batched `SPHttpClient` requests. Field Customizers should include ARIA labels and meet WCAG 2.1 AA.
- **PnP templates**: Use the 2023/01 PnP provisioning schema. Include content type bindings and view definitions.
- **PowerShell**: Use `[CmdletBinding(SupportsShouldProcess)]` on all scripts. Use `PnP.PowerShell` cmdlets, not legacy CSOM. Follow the `Write-StatusLine` convention.

## Submitting Changes

1. Ensure TypeScript compiles without errors (`gulp build`).
2. Ensure PowerShell scripts pass `Invoke-ScriptAnalyzer`.
3. Ensure PnP XML templates are valid against the schema.
4. Include screenshots or HTML mockups for UI changes.
5. Open a Pull Request with a clear description of what changed and why.
