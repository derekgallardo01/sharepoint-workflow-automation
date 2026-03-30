# Item Lifecycle Management

This diagram shows the **Lifecycle Management** Power Automate flow, which runs on a daily schedule to handle archival, reminders, and overdue status updates across all project lists.

```mermaid
flowchart TD
    A([Daily Trigger - 6:00 AM]) --> B[Initialize Audit Log Array]

    B --> C[Query All Active Lists]

    C --> D[Branch 1:<br/>Archive Old Items]
    C --> E[Branch 2:<br/>Due Date Reminders]
    C --> F[Branch 3:<br/>Overdue Detection]

    %% Branch 1: Archive
    D --> D1["Filter: Status = 'Completed'<br/>AND CompletedDate < Today - 90 days"]
    D1 --> D2{Items Found?}
    D2 -->|Yes| D3[For Each Item]
    D2 -->|No| D7[Log: No items to archive]
    D3 --> D4[Copy Item to Archive List]
    D4 --> D5[Delete from Source List]
    D5 --> D6["Log: Archived {Title}<br/>from {ListName}"]
    D6 --> D3

    %% Branch 2: Reminders
    E --> E1["Filter: Status != 'Completed'<br/>AND DueDate = Tomorrow"]
    E1 --> E2{Items Found?}
    E2 -->|Yes| E3[For Each Item]
    E2 -->|No| E7[Log: No reminders needed]
    E3 --> E4[Get Assigned User Email]
    E4 --> E5["Send Reminder Email:<br/>Subject: Due Tomorrow - {Title}<br/>Body: Item details + direct link"]
    E5 --> E6["Log: Reminder sent for<br/>{Title} to {AssignedTo}"]
    E6 --> E3

    %% Branch 3: Overdue
    F --> F1["Filter: Status NOT IN<br/>('Completed','Archived','Overdue')<br/>AND DueDate < Today"]
    F1 --> F2{Items Found?}
    F2 -->|Yes| F3[For Each Item]
    F2 -->|No| F7[Log: No overdue items]
    F3 --> F4["Update Status to 'Overdue'"]
    F4 --> F5[Send Overdue Alert to<br/>Assigned User + Manager]
    F5 --> F6["Log: Marked overdue -<br/>{Title}, was {OldStatus}"]
    F6 --> F3

    %% Converge
    D7 --> G[Compile Audit Summary]
    D6 --> G
    E7 --> G
    E6 --> G
    F7 --> G
    F6 --> G

    G --> H["Write Audit Entry:<br/>Date, ItemsArchived,<br/>RemindersSent, ItemsMarkedOverdue"]
    H --> I([End])

    %% Styling
    style A fill:#0078d4,stroke:#005a9e,color:#fff
    style D fill:#6b69d6,stroke:#4a48b5,color:#fff
    style E fill:#6b69d6,stroke:#4a48b5,color:#fff
    style F fill:#6b69d6,stroke:#4a48b5,color:#fff
    style D4 fill:#107c10,stroke:#0b5a08,color:#fff
    style D5 fill:#d83b01,stroke:#a52c00,color:#fff
    style E5 fill:#0078d4,stroke:#005a9e,color:#fff
    style F4 fill:#d83b01,stroke:#a52c00,color:#fff
    style F5 fill:#d83b01,stroke:#a52c00,color:#fff
    style H fill:#107c10,stroke:#0b5a08,color:#fff
    style I fill:#333,stroke:#222,color:#fff
```

## Daily Processing Summary

| Branch | Filter Criteria | Action | Notification |
|--------|----------------|--------|--------------|
| **Archive** | Status = Completed AND CompletedDate > 90 days ago | Copy to Archive list, delete from source | None (logged only) |
| **Reminders** | Status != Completed AND DueDate = Tomorrow | Send email reminder | Email to assigned user |
| **Overdue** | Status not terminal AND DueDate < Today | Update status to "Overdue" | Email to assigned user + their manager |

## Schedule

- **Trigger**: Daily recurrence at 6:00 AM (site timezone)
- **Estimated duration**: 2-5 minutes depending on item count
- **Retry policy**: 3 retries with exponential backoff on transient failures

## Audit Trail

Every run produces an audit entry in the `WorkflowAuditLog` list with:

| Field | Description |
|-------|-------------|
| `RunDate` | Timestamp of the flow run |
| `ItemsArchived` | Count of items moved to archive |
| `RemindersSent` | Count of reminder emails sent |
| `ItemsMarkedOverdue` | Count of items updated to Overdue status |
| `Errors` | Any errors encountered during processing |
| `Duration` | Total flow run time in seconds |
