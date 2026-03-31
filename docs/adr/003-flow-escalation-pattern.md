# ADR 003: Parallel Branch Escalation Pattern for Approvals

**Status:** Accepted
**Date:** 2026-01-22
**Decision Makers:** Solution Architect, Power Platform Lead

## Context

The Multi-Stage Approval flow requires escalation logic: if an approver does not respond within a defined period, the system must send reminders and eventually escalate to a higher authority. Two primary patterns exist in Power Automate:

1. **Serial approach**: Set a timeout on the "Start and wait for an approval" action, then handle the timeout in a subsequent step.
2. **Parallel branch approach**: Run the approval action alongside a parallel branch that independently monitors elapsed time and triggers reminders/escalation.

## Decision

We will use the **parallel branch pattern** where the approval action and the escalation timer run as concurrent branches within a `Parallel` scope.

## Design

```
[Scope: Approval with Escalation]
├── Branch 1: Approval
│   └── Start and wait for an approval (no timeout)
│       └── On response → terminate parallel branch, continue flow
│
├── Branch 2: Reminder (3-day timer)
│   └── Delay 3 days
│       └── Condition: Has approval completed?
│           ├── Yes → Do nothing (branch terminates)
│           └── No → Send reminder email to approver
│               └── Post reminder to Teams
│
└── Branch 3: Escalation (5-day timer)
    └── Delay 5 days
        └── Condition: Has approval completed?
            ├── Yes → Do nothing (branch terminates)
            └── No → Cancel original approval
                └── Create new approval for skip-level manager
                    └── Update item: "Escalated to [VP Name]"
```

### Synchronization

The scope is configured with `runAfter` settings so that the flow continues after Branch 1 completes (either by response or cancellation from Branch 3). Branches 2 and 3 check a flow variable `approvalCompleted` (set by Branch 1 on response) to avoid sending unnecessary reminders.

## Rationale

### Why Not Serial (Timeout on Approval Action)

Power Automate's "Start and wait for an approval" supports a `timeout` property (ISO 8601 duration, e.g., `P3D`). The serial approach would be:

```
Start and wait for approval (timeout: P5D)
├── If Outcome = "Approve" → continue
├── If Outcome = "Reject" → handle rejection
└── If timed out → escalate
```

**Problems with serial timeout:**

1. **No intermediate reminders**: The timeout is all-or-nothing. You get either a response or a timeout after 5 days. There is no built-in way to send a reminder at day 3 while still waiting for the original approval.
2. **Blocking execution**: The flow run is completely blocked during the wait period. You cannot update the SharePoint item with "Awaiting response" status or log any intermediate audit entries.
3. **Timeout granularity**: If you set timeout to 3 days for a reminder, you lose the original approval action and must create a new one, resetting the approver's context in Teams/Outlook.
4. **No escalation flexibility**: Escalation (skip-level approval) requires canceling the original approval and creating a new one. With serial timeout, you handle this after the wait completes, adding the full timeout duration before any escalation occurs.

### Advantages of Parallel Branch

- **Reminder at day 3, escalation at day 5**: Both timers run independently alongside the approval. The approver can respond at any point during either timer.
- **Non-blocking status updates**: Branch 2 can update the SharePoint item status to "Reminder Sent" at day 3, providing visibility in the list view.
- **Graceful cancellation**: If the approver responds before either timer fires, the flow variable `approvalCompleted = true` causes the timer branches to no-op and terminate.
- **Audit trail**: Each branch can log its actions to the Workflow Audit Log independently (reminder sent, escalation triggered, etc.).
- **Composable**: Additional branches can be added (e.g., a 1-day "soft nudge" via Teams) without restructuring the flow.

### Trade-offs

| Aspect | Parallel Branch | Serial Timeout |
|---|---|---|
| Complexity | Higher (3 branches, shared variable) | Lower (single action with timeout) |
| Reminder capability | Yes (independent timer) | No (all-or-nothing) |
| Flow run duration | Same (waits for approval either way) | Same |
| API actions consumed | Slightly more (parallel branches count) | Fewer |
| Debuggability | Each branch visible in run history | Single linear path |
| Escalation flexibility | Cancel + re-create approval mid-wait | Must wait for full timeout |

The added complexity is justified by the requirement for intermediate reminders and flexible escalation. The parallel pattern is a well-established Power Automate design pattern documented by Microsoft MVP community.

## Consequences

- The Multi-Stage Approval flow uses `Parallel` scopes at both approval stages (Manager and Department Head).
- A flow-level variable `stage1Completed` / `stage2Completed` (boolean) synchronizes branches.
- The Workflow Audit Log receives entries for: approval sent, reminder sent (3 days), escalation triggered (5 days), approval response received.
- Flow run history shows all three branches, making it clear which path executed.
- The reminder and escalation intervals (3 days, 5 days) are defined as flow-level parameters, making them configurable without editing actions.

## Alternatives Considered

| Alternative | Reason for Rejection |
|---|---|
| Serial timeout on approval action | No intermediate reminders; blocking; inflexible escalation |
| Separate scheduled flow for reminders | Requires cross-flow coordination; hard to cancel when approval completes |
| Power Automate "Do until" loop with delay | Polling pattern; consumes API actions on each loop iteration; harder to reason about |
| Azure Logic Apps with durable functions | Over-engineered for this scenario; requires Azure subscription and separate deployment |
