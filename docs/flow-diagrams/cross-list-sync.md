# Cross-List Data Synchronization

This diagram shows the **Cross-List Sync** Power Automate flow, which implements a hub-and-spoke synchronization pattern. When the hub list is modified, changes propagate to all matching spoke list items with hub-wins conflict resolution.

```mermaid
flowchart TD
    A([Hub List Item Modified]) --> B[Read Modified Item Fields]
    B --> C[Query Spoke Lists for<br/>Matching Items by LookupId]

    C --> D{Matching Spoke<br/>Items Found?}

    D -->|No| E[Log: No Matching Items<br/>to Audit List]
    D -->|Yes| F[Initialize Change Counter = 0]

    F --> G[For Each Spoke Item]

    G --> H[Compare Hub Fields<br/>vs Spoke Fields]

    H --> I{Fields<br/>Changed?}

    I -->|No| J[Skip - Already in Sync]
    I -->|Yes| K{Conflict<br/>Detected?}

    K -->|Yes - Hub Wins| L[Overwrite Spoke with Hub Data]
    K -->|No Conflict| L

    L --> M[Update Spoke Item<br/>via SharePoint REST]
    M --> N[Increment Change Counter]

    J --> O{More Spoke Items?}
    N --> O

    O -->|Yes| G
    O -->|No| P[Log Sync Results<br/>to Audit List]

    P --> Q["Audit Entry:<br/>HubItemId, SpokeList,<br/>ItemsChecked, ItemsUpdated,<br/>Timestamp, ConflictsResolved"]

    E --> R([End])
    Q --> R

    %% Conflict Resolution Detail
    K -.-> K1["Conflict Resolution Policy:<br/>• Hub data ALWAYS wins<br/>• Spoke modifications are overwritten<br/>• Original spoke values logged<br/>  before overwrite for rollback"]

    %% Styling
    style A fill:#0078d4,stroke:#005a9e,color:#fff
    style L fill:#107c10,stroke:#0b5a08,color:#fff
    style M fill:#107c10,stroke:#0b5a08,color:#fff
    style E fill:#808080,stroke:#606060,color:#fff
    style J fill:#808080,stroke:#606060,color:#fff
    style K fill:#d83b01,stroke:#a52c00,color:#fff
    style Q fill:#f5f5f5,stroke:#ccc,color:#333
    style K1 fill:#fff3cd,stroke:#d83b01,color:#333
    style R fill:#333,stroke:#222,color:#fff
```

## Synchronization Rules

| Rule | Behavior |
|------|----------|
| **Direction** | One-way: Hub to Spoke(s) |
| **Trigger** | Any field modification on a Hub list item |
| **Matching** | Spoke items linked via `HubItemLookupId` column |
| **Conflict Resolution** | Hub data always wins; spoke values overwritten |
| **Audit Trail** | Every sync operation logged with before/after values |
| **Rollback** | Previous spoke values stored in audit log for manual rollback |

## Synchronized Fields

| Hub Field | Spoke Field | Sync Behavior |
|-----------|-------------|---------------|
| Title | Title | Direct copy |
| Status | Status | Direct copy |
| Priority | Priority | Direct copy |
| AssignedTo | AssignedTo | User lookup copy |
| DueDate | DueDate | Direct copy |
| Category | Category | Direct copy |

## Performance Notes

- The flow uses batched REST calls (`$batch` endpoint) to minimize API calls when updating multiple spoke items.
- A maximum of 50 spoke items are processed per run to stay within Power Automate API limits.
- If more than 50 matches exist, subsequent items are queued for the next trigger cycle.
