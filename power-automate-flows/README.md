# Power Automate Flow Definitions

This directory contains exportable Power Automate flow definitions for the SharePoint Workflow Automation solution. Each flow is defined as a Logic Apps workflow JSON schema, compatible with Power Automate import.

## Flows Overview

| Flow | File | Trigger | Purpose |
|------|------|---------|---------|
| Multi-Stage Approval | `multi-stage-approval.json` | Item created/modified | Routes documents through manager and department head approval with escalation |
| Conditional Notifications | `conditional-notifications.json` | Item created | Routes notifications by priority: Teams + email, Teams only, or weekly digest |
| Weekly Status Report | `scheduled-report.json` | Recurrence (Mon 8 AM) | Generates HTML/PDF report of active projects and emails stakeholders |
| Cross-List Sync | `cross-list-sync.json` | Item modified | Synchronizes hub list changes to spoke lists with conflict resolution |
| Lifecycle Management | `lifecycle-management.json` | Recurrence (daily 6 AM) | Archives old items, sends due-date reminders, marks overdue items |

## Prerequisites

Before importing these flows, ensure the following connectors are available in your Power Automate environment:

- **SharePoint** (standard connector)
- **Office 365 Outlook** (standard connector)
- **Office 365 Users** (standard connector)
- **Approvals** (standard connector)
- **Microsoft Teams** (standard connector)
- **Encodian** (premium connector, used for HTML-to-PDF in the report flow -- can be substituted)

## Import Instructions

### Method 1: Power Automate Portal

1. Navigate to [Power Automate](https://make.powerautomate.com)
2. Select **My flows** from the left navigation
3. Click **Import** > **Import Package (Legacy)** or use **Import from JSON**
4. Upload the desired `.json` file
5. Configure the required connections when prompted
6. Update the flow parameters (site URL, list IDs, etc.) to match your environment

### Method 2: Power Automate CLI

```bash
# Install the Power Platform CLI if not already installed
pac install latest

# Authenticate
pac auth create --environment https://contoso.crm.dynamics.com

# Import flow
pac solution import --path ./multi-stage-approval.json
```

### Method 3: Power Automate Management Connector (programmatic)

Use the Power Automate Management connector within another flow to deploy these definitions programmatically.

## Configuration

After importing each flow, update the following parameters:

### Common Parameters (all flows)

| Parameter | Description | Example |
|-----------|-------------|---------|
| `siteUrl` | SharePoint site URL | `https://contoso.sharepoint.com/sites/workflow-demo` |

### Flow-Specific Parameters

**Multi-Stage Approval**
| Parameter | Description |
|-----------|-------------|
| `approvalListId` | GUID of the Document Approval Queue list |
| `escalationDaysReminder` | Days before sending a reminder (default: 3) |
| `escalationDaysAutoEscalate` | Days before auto-escalation (default: 5) |

**Conditional Notifications**
| Parameter | Description |
|-----------|-------------|
| `projectListId` | GUID of the Project Tracker list |
| `teamsChannelId` | Target Teams channel for notifications |
| `teamsGroupId` | Teams group (team) ID |

**Weekly Status Report**
| Parameter | Description |
|-----------|-------------|
| `projectListId` | GUID of the Project Tracker list |
| `reportLibraryPath` | SharePoint library path for saving reports |
| `reportRecipients` | Semicolon-separated email addresses |

**Cross-List Sync**
| Parameter | Description |
|-----------|-------------|
| `hubListId` | GUID of the hub (source) list |
| `spokeListIds` | Array of spoke list GUIDs |
| `auditListId` | GUID of the sync audit log list |
| `lookupFieldName` | Internal name of the lookup field on spoke lists |
| `syncFields` | Array of field internal names to synchronize |

**Lifecycle Management**
| Parameter | Description |
|-----------|-------------|
| `projectListId` | GUID of the Project Tracker list |
| `archiveListId` | GUID of the archive list |
| `auditListId` | GUID of the audit log list |
| `archiveAfterDays` | Days after completion before archiving (default: 90) |

## Customization Notes

- **HTML-to-PDF**: The weekly report flow uses the Encodian connector for PDF generation. If you do not have an Encodian license, you can substitute with the **Muhimbi** connector, **Adobe PDF Services** connector, or simply send the HTML email directly without PDF generation.
- **Escalation Timers**: The multi-stage approval flow uses parallel branches for escalation. Adjust the `escalationDaysReminder` and `escalationDaysAutoEscalate` parameters to match your organization's SLA requirements.
- **Conflict Resolution**: The cross-list sync flow uses a "hub wins" strategy. If you need "last write wins" or manual conflict resolution, modify the `Check_Conflict_Hub_Wins` condition.

## Troubleshooting

| Issue | Solution |
|-------|----------|
| Flow fails on SharePoint connector | Verify the site URL and list IDs are correct; check that the connection account has Contribute permissions |
| Approval not received | Ensure the Approvals connector is configured and the approver's mailbox is active |
| Teams notification fails | Verify the Teams group ID and channel ID; ensure the connection account is a member of the team |
| PDF generation fails | Check the Encodian connector license; consider switching to an alternative PDF connector |
| Sync flow loops | Ensure the spoke list update does not re-trigger the hub list webhook; add a condition to check `Editor` |
