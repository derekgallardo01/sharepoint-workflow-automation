# Smart Notification Routing

This diagram shows how the **Conditional Notifications** Power Automate flow routes alerts based on item priority. High-priority items trigger immediate multi-channel notifications, while low-priority items are batched into a weekly digest.

```mermaid
flowchart TD
    A([New Item Created in List]) --> B[Read Item Priority Field]
    B --> C{Priority Level?}

    C -->|High| D[Teams Channel Notification]
    C -->|Medium| G[Teams Notification Only]
    C -->|Low| J[Add to Weekly Digest Queue]

    %% High Priority Path
    D --> E[Email to Assigned Manager]
    E --> F["Format: 🔴 URGENT<br/>Title: {ItemTitle}<br/>Priority: High<br/>Due: {DueDate}<br/>Action Required Immediately"]
    F --> L([End])

    %% Medium Priority Path
    G --> H["Format: 🟡 ATTENTION<br/>Title: {ItemTitle}<br/>Priority: Medium<br/>Due: {DueDate}<br/>Please review at your convenience"]
    H --> L

    %% Low Priority Path
    J --> K["Format: Added to Digest<br/>Title: {ItemTitle}<br/>Priority: Low<br/>Will be included in next<br/>Monday weekly summary"]
    K --> L

    %% Teams Notification Details
    D -.-> D1["Channel: #project-alerts<br/>Adaptive Card with action buttons:<br/>• View Item<br/>• Approve<br/>• Assign to Me"]
    E -.-> E1["To: Manager email<br/>Subject: Action Required - {ItemTitle}<br/>Body: HTML with item details,<br/>direct link, and response buttons"]
    G -.-> G1["Channel: #project-updates<br/>Simple message card:<br/>• Item title and link<br/>• Priority badge<br/>• Assigned to"]
    J -.-> J1["Adds row to DigestQueue list<br/>Fields: ItemId, Title, Priority,<br/>CreatedDate, ListName<br/>Processed by Weekly Report flow"]

    %% Styling
    style A fill:#0078d4,stroke:#005a9e,color:#fff
    style D fill:#a80000,stroke:#750000,color:#fff
    style E fill:#a80000,stroke:#750000,color:#fff
    style F fill:#a80000,stroke:#750000,color:#fff
    style G fill:#d83b01,stroke:#a52c00,color:#fff
    style H fill:#d83b01,stroke:#a52c00,color:#fff
    style J fill:#107c10,stroke:#0b5a08,color:#fff
    style K fill:#107c10,stroke:#0b5a08,color:#fff
    style L fill:#333,stroke:#222,color:#fff
    style D1 fill:#f5f5f5,stroke:#ccc,color:#333
    style E1 fill:#f5f5f5,stroke:#ccc,color:#333
    style G1 fill:#f5f5f5,stroke:#ccc,color:#333
    style J1 fill:#f5f5f5,stroke:#ccc,color:#333
```

## Notification Formats

| Priority | Channel | Format | Urgency |
|----------|---------|--------|---------|
| **High** | Teams Channel + Manager Email | Adaptive Card with action buttons + HTML email | Immediate response required |
| **Medium** | Teams Channel only | Simple message card with item link | Review at convenience |
| **Low** | Weekly Digest Queue | Batched into Monday summary report | Informational only |

## Configuration

The priority thresholds and notification targets are configured in the flow's environment variables:

- `TeamsChannelId_Alerts` -- Channel for high-priority notifications
- `TeamsChannelId_Updates` -- Channel for medium-priority notifications
- `DigestQueueListName` -- Name of the list used for weekly digest batching
- `ManagerLookupField` -- Field name used to resolve the assigned manager
