# Multi-Stage Document Approval Flow

This diagram illustrates the complete two-stage approval routing used by the **Multi-Stage Approval** Power Automate flow. Documents pass through manager review, then department head review, with automatic escalation if approvers are unresponsive.

```mermaid
flowchart TD
    A([Document Submitted]) --> B[Get Submitter's Manager]
    B --> C{Manager Approval?}

    C -->|Approved| D[Get Department Head]
    C -->|Rejected| E[Mark Item as Rejected]

    D --> F{Dept Head Approval?}

    F -->|Approved| G[Mark Item as Approved]
    F -->|Rejected| E

    G --> H[Send Approval Notification to Submitter]
    E --> I[Send Rejection Email with Reason]

    H --> J([End])
    I --> J

    %% Escalation / Reminder Path
    C -.->|No Response| K{3 Days Elapsed?}
    K -->|Yes| L[Send Reminder to Manager]
    L --> M{5 Days Elapsed?}
    M -->|Yes| N[Auto-Escalate to Department Head]
    M -->|No| C
    N --> F

    F -.->|No Response| O{3 Days Elapsed?}
    O -->|Yes| P[Send Reminder to Dept Head]
    P --> Q{5 Days Elapsed?}
    Q -->|Yes| R[Auto-Escalate to VP / Skip Level]
    Q -->|No| F
    R --> S[Mark Escalated & Approved]
    S --> H

    %% Styling
    style A fill:#0078d4,stroke:#005a9e,color:#fff
    style G fill:#107c10,stroke:#0b5a08,color:#fff
    style H fill:#107c10,stroke:#0b5a08,color:#fff
    style S fill:#107c10,stroke:#0b5a08,color:#fff
    style E fill:#a80000,stroke:#750000,color:#fff
    style I fill:#a80000,stroke:#750000,color:#fff
    style L fill:#d83b01,stroke:#a52c00,color:#fff
    style N fill:#d83b01,stroke:#a52c00,color:#fff
    style P fill:#d83b01,stroke:#a52c00,color:#fff
    style R fill:#d83b01,stroke:#a52c00,color:#fff
    style J fill:#333,stroke:#222,color:#fff
```

## Flow Summary

| Stage | Approver | Timeout | Escalation Target |
|-------|----------|---------|-------------------|
| Stage 1 | Submitter's Manager | 3-day reminder, 5-day auto-escalate | Department Head |
| Stage 2 | Department Head | 3-day reminder, 5-day auto-escalate | VP / Skip Level |

## Key Behaviors

- **Rejection at any stage** terminates the flow and sends a rejection email with the approver's comments.
- **Approval at both stages** marks the document as Approved and notifies the original submitter.
- **Escalation** bypasses the unresponsive approver and moves the request to the next level in the hierarchy.
