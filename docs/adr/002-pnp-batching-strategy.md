# ADR 002: Use PnP Batched Operations for Bulk Updates

**Status:** Accepted
**Date:** 2026-01-18
**Decision Makers:** Solution Architect, SharePoint Lead

## Context

The Bulk Approve action may update 10-50+ items in a single operation. SharePoint Online REST API enforces:

- **Request throttling**: HTTP 429 responses when exceeding ~600 requests/minute per app.
- **Per-request latency**: Each individual `PATCH` to `/_api/web/lists/getByTitle(...)/items(id)` takes 200-400ms round-trip.
- **User experience**: Sequential updates of 50 items at 300ms each = 15 seconds of spinner time.

We need a strategy that minimizes API calls, respects throttling limits, and provides acceptable UX for bulk status changes.

## Decision

We will use **PnP/SP `createBatch()` batched operations** to combine multiple item updates into a single HTTP `$batch` request, with a configurable batch size and retry strategy.

## Rationale

### How SharePoint $batch Works

SharePoint REST API supports OData `$batch` requests, which bundle multiple operations into a single HTTP POST to `/_api/$batch`. PnP/SP abstracts this with:

```typescript
const [batchedSP, execute] = sp.batched();

for (const item of selectedItems) {
  batchedSP.web.lists
    .getByTitle(listTitle)
    .items.getById(item.id)
    .update({ Status: "Approved", ApprovedBy: currentUser, ApprovedDate: now });
}

await execute();
```

### Performance Impact

| Approach | Items | API Calls | Est. Duration | Throttle Risk |
|---|---|---|---|---|
| Sequential REST calls | 50 | 50 | ~15s | High |
| PnP batched (single batch) | 50 | 1 | ~1.2s | None |
| PnP batched (chunks of 20) | 50 | 3 | ~2.5s | Negligible |

**Result**: Batch reduces N calls to ceil(N / batchSize) calls. For typical use (10-20 items), this is always 1 call.

### Batch Size Selection

SharePoint $batch has an undocumented limit of ~100 operations per batch. We use a conservative **batch size of 20** as the default:

- **20 items per batch**: Well within the $batch limit, leaves headroom for complex operations that may expand into multiple sub-requests (e.g., update item + break permission inheritance).
- Configurable via `BatchOptions.batchSize` for environments with different throttling profiles.

### Trade-offs: All-or-Nothing vs. Individual Error Handling

**$batch semantics**: SharePoint processes all operations in a batch but does **not** roll back on partial failure. Each operation in the batch returns its own HTTP status code. A batch of 20 updates may have 18 succeed and 2 fail (e.g., due to version conflict or missing field).

This is actually preferable for our use case:
- **No artificial rollback**: If 18 of 20 items updated successfully, we want to keep those 18. Rolling back all 20 because of 2 failures wastes user work.
- **Granular error reporting**: The `BatchOperationService` parses individual response codes from the batch response and reports per-item success/failure to the `ProgressPanel`.
- **Retry failed items only**: Failed items are collected and can be retried in a subsequent batch without re-processing the already-succeeded items.

### Retry Strategy

Failed items are retried with exponential backoff:

| Attempt | Delay | Behavior |
|---|---|---|
| 1 | 0ms | Immediate (initial batch) |
| 2 | 1000ms | Retry failed items only |
| 3 | 3000ms | Retry remaining failures |
| 4+ | N/A | Report as permanent failure |

If the failure is HTTP 429 (throttled), we honor the `Retry-After` header from SharePoint before the next batch.

## Consequences

- All bulk update operations go through `BatchOperationService.batchProcess()`.
- The service reports progress via callback: `onProgress({ completed, total, succeeded, failed, currentItem })`.
- Failed items surface in the ProgressPanel with error details and per-item retry buttons.
- Audit logging captures the batch operation as a single logical action with individual item outcomes.
- Batch size is configurable but defaults to 20. This should be tuned if Microsoft changes $batch limits.

## Alternatives Considered

| Alternative | Reason for Rejection |
|---|---|
| Sequential REST calls | Too slow for 10+ items; high throttle risk |
| Microsoft Graph batch ($batch on Graph) | Graph does not support SharePoint list item updates in batch as of 2026 |
| Power Automate "Apply to each" | Runs server-side but no real-time progress feedback in SPFx; adds flow run latency |
| PnP PowerShell Batch-PnPListItem | Server-side only; not available in browser-based SPFx context |
