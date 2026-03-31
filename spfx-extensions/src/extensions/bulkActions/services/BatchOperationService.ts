import { SPFI } from "@pnp/sp";
import "@pnp/sp/batching";
import "@pnp/sp/items";
import "@pnp/sp/lists";
import "@pnp/sp/webs";

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/**
 * Configuration options for batch processing operations.
 */
export interface BatchOptions {
  /** Maximum items per $batch request. Default: 20. Max recommended: 50. */
  batchSize: number;

  /** Maximum parallel batches in flight. Default: 5. */
  concurrency: number;

  /** Maximum retry attempts for failed items. Default: 3. */
  maxRetries: number;

  /** Base delay (ms) for exponential backoff between retries. Default: 1000. */
  retryBaseDelay: number;

  /** Whether to respect SharePoint 429 Retry-After headers. Default: true. */
  respectThrottling: boolean;

  /** Abort signal for cancellation support. */
  cancellationToken?: AbortSignal;
}

/**
 * Status of an individual item within a batch operation.
 */
export enum ItemOperationStatus {
  Pending = "Pending",
  Processing = "Processing",
  Succeeded = "Succeeded",
  Failed = "Failed",
  Retrying = "Retrying",
  Cancelled = "Cancelled",
}

/**
 * Result of processing a single item.
 */
export interface ItemResult<T> {
  item: T;
  status: ItemOperationStatus;
  error?: string;
  attempts: number;
  duration: number;
}

/**
 * Progress report emitted during batch processing.
 */
export interface BatchProgress<T> {
  /** Total items in the operation. */
  total: number;
  /** Items completed (succeeded + failed after all retries). */
  completed: number;
  /** Items that succeeded. */
  succeeded: number;
  /** Items that failed after all retries. */
  failed: number;
  /** Items currently being processed. */
  inProgress: number;
  /** The item currently being reported on (if applicable). */
  currentItem?: T;
  /** Current status of the reported item. */
  currentStatus?: ItemOperationStatus;
  /** Per-item results accumulated so far. */
  results: ItemResult<T>[];
  /** Elapsed time in milliseconds. */
  elapsed: number;
}

/**
 * Final result of a completed batch operation.
 */
export interface BatchResult<T> {
  /** Whether all items succeeded. */
  success: boolean;
  /** Total items processed. */
  total: number;
  /** Number of items that succeeded. */
  succeeded: number;
  /** Number of items that failed after all retries. */
  failed: number;
  /** Number of items cancelled before processing. */
  cancelled: number;
  /** Per-item results. */
  results: ItemResult<T>[];
  /** Total elapsed time in milliseconds. */
  duration: number;
}

/**
 * Entry in the operation audit log.
 */
export interface AuditEntry {
  timestamp: Date;
  operation: string;
  itemId: string | number;
  itemTitle: string;
  user: string;
  previousValue?: string;
  newValue?: string;
  status: "Success" | "Failed" | "Cancelled";
  error?: string;
  batchId: string;
}

// ---------------------------------------------------------------------------
// Defaults
// ---------------------------------------------------------------------------

const DEFAULT_OPTIONS: BatchOptions = {
  batchSize: 20,
  concurrency: 5,
  maxRetries: 3,
  retryBaseDelay: 1000,
  respectThrottling: true,
};

// ---------------------------------------------------------------------------
// Service
// ---------------------------------------------------------------------------

/**
 * Advanced batch operation service for SharePoint list item updates.
 *
 * Features:
 * - Generic batch processor with configurable concurrency
 * - Progress callback with detailed per-item status
 * - Retry failed items with exponential backoff
 * - Cancellation token support (AbortSignal)
 * - Operation audit log generation
 * - SharePoint throttling (HTTP 429) detection and backoff
 */
export class BatchOperationService {
  private readonly _sp: SPFI;
  private readonly _auditLog: AuditEntry[] = [];

  constructor(sp: SPFI) {
    this._sp = sp;
  }

  // -----------------------------------------------------------------------
  // Public API
  // -----------------------------------------------------------------------

  /**
   * Process a collection of items in batched, concurrent operations with
   * retry logic, progress reporting, and cancellation support.
   *
   * @param items      - The items to process.
   * @param processor  - An async function that performs the operation on a single item.
   *                     Throw an error to indicate failure (the item will be retried).
   * @param options    - Batch configuration (merged with defaults).
   * @param onProgress - Optional callback invoked after each item completes or fails.
   * @returns A promise resolving to the final BatchResult.
   *
   * @example
   * ```typescript
   * const service = new BatchOperationService(sp);
   *
   * const result = await service.batchProcess(
   *   selectedItems,
   *   async (item) => {
   *     await sp.web.lists.getByTitle("Projects").items.getById(item.Id).update({
   *       Status: "Approved",
   *       ApprovedDate: new Date().toISOString(),
   *     });
   *   },
   *   { batchSize: 20, concurrency: 5, maxRetries: 3 },
   *   (progress) => {
   *     console.log(`${progress.completed}/${progress.total} complete`);
   *   }
   * );
   * ```
   */
  public async batchProcess<T>(
    items: T[],
    processor: (item: T) => Promise<void>,
    options?: Partial<BatchOptions>,
    onProgress?: (progress: BatchProgress<T>) => void
  ): Promise<BatchResult<T>> {
    const opts: BatchOptions = { ...DEFAULT_OPTIONS, ...options };
    const startTime = Date.now();
    const batchId = this._generateBatchId();

    // Initialize per-item tracking
    const results: ItemResult<T>[] = items.map((item) => ({
      item,
      status: ItemOperationStatus.Pending,
      attempts: 0,
      duration: 0,
    }));

    let succeeded = 0;
    let failed = 0;
    let cancelled = 0;

    // Helper: emit progress
    const emitProgress = (currentItem?: T, currentStatus?: ItemOperationStatus): void => {
      if (!onProgress) return;
      const completed = succeeded + failed + cancelled;
      onProgress({
        total: items.length,
        completed,
        succeeded,
        failed,
        inProgress: items.length - completed,
        currentItem,
        currentStatus,
        results: [...results],
        elapsed: Date.now() - startTime,
      });
    };

    // Split items into chunks of batchSize
    const pendingIndices = items.map((_, i) => i);
    let retryRound = 0;

    while (pendingIndices.length > 0 && retryRound <= opts.maxRetries) {
      // Check cancellation
      if (opts.cancellationToken?.aborted) {
        for (const idx of pendingIndices) {
          results[idx].status = ItemOperationStatus.Cancelled;
          cancelled++;
        }
        pendingIndices.length = 0;
        emitProgress();
        break;
      }

      // Apply backoff delay for retries (not the first round)
      if (retryRound > 0) {
        const delay = this._calculateBackoff(retryRound, opts.retryBaseDelay);
        await this._delay(delay, opts.cancellationToken);
      }

      // Chunk the pending indices into batches
      const chunks = this._chunk(pendingIndices, opts.batchSize);

      // Process chunks with bounded concurrency
      const failedInThisRound: number[] = [];

      for (const concurrentBatch of this._chunk(chunks, opts.concurrency)) {
        // Check cancellation before each concurrent batch
        if (opts.cancellationToken?.aborted) break;

        const batchPromises = concurrentBatch.map(async (chunk) => {
          for (const idx of chunk) {
            if (opts.cancellationToken?.aborted) {
              results[idx].status = ItemOperationStatus.Cancelled;
              cancelled++;
              emitProgress(results[idx].item, ItemOperationStatus.Cancelled);
              continue;
            }

            const itemStart = Date.now();
            results[idx].status =
              retryRound === 0 ? ItemOperationStatus.Processing : ItemOperationStatus.Retrying;
            results[idx].attempts++;
            emitProgress(results[idx].item, results[idx].status);

            try {
              await processor(results[idx].item);
              results[idx].status = ItemOperationStatus.Succeeded;
              results[idx].duration = Date.now() - itemStart;
              succeeded++;
              emitProgress(results[idx].item, ItemOperationStatus.Succeeded);
            } catch (err) {
              results[idx].duration = Date.now() - itemStart;

              // Check for HTTP 429 throttling
              if (opts.respectThrottling && this._isThrottled(err)) {
                const retryAfter = this._getRetryAfterMs(err);
                await this._delay(retryAfter, opts.cancellationToken);
              }

              if (results[idx].attempts >= opts.maxRetries) {
                results[idx].status = ItemOperationStatus.Failed;
                results[idx].error = this._extractErrorMessage(err);
                failed++;
                emitProgress(results[idx].item, ItemOperationStatus.Failed);
              } else {
                failedInThisRound.push(idx);
              }
            }
          }
        });

        await Promise.all(batchPromises);
      }

      // Prepare next retry round with only the items that failed but have retries left
      pendingIndices.length = 0;
      pendingIndices.push(...failedInThisRound);
      retryRound++;
    }

    // Any items still pending after all retries are marked as failed
    for (const idx of pendingIndices) {
      if (results[idx].status !== ItemOperationStatus.Succeeded &&
          results[idx].status !== ItemOperationStatus.Cancelled) {
        results[idx].status = ItemOperationStatus.Failed;
        results[idx].error = results[idx].error || "Max retries exceeded";
        failed++;
      }
    }

    const totalDuration = Date.now() - startTime;

    return {
      success: failed === 0 && cancelled === 0,
      total: items.length,
      succeeded,
      failed,
      cancelled,
      results,
      duration: totalDuration,
    };
  }

  /**
   * Retry specific failed items from a previous batch result.
   * Useful for the "Retry Failed" button in the ProgressPanel.
   */
  public async retryFailed<T>(
    previousResult: BatchResult<T>,
    processor: (item: T) => Promise<void>,
    options?: Partial<BatchOptions>,
    onProgress?: (progress: BatchProgress<T>) => void
  ): Promise<BatchResult<T>> {
    const failedItems = previousResult.results
      .filter((r) => r.status === ItemOperationStatus.Failed)
      .map((r) => r.item);

    if (failedItems.length === 0) {
      return {
        success: true,
        total: 0,
        succeeded: 0,
        failed: 0,
        cancelled: 0,
        results: [],
        duration: 0,
      };
    }

    return this.batchProcess(failedItems, processor, options, onProgress);
  }

  /**
   * Generate audit entries for a completed batch operation.
   * Writes entries to the internal audit log and returns them for
   * persistence to the Workflow Audit Log list.
   */
  public generateAuditEntries<T extends { Id?: number; Title?: string }>(
    batchResult: BatchResult<T>,
    operation: string,
    user: string,
    fieldChanges?: { field: string; previousValue?: string; newValue: string }
  ): AuditEntry[] {
    const batchId = this._generateBatchId();
    const entries: AuditEntry[] = batchResult.results.map((r) => ({
      timestamp: new Date(),
      operation,
      itemId: (r.item as T).Id ?? 0,
      itemTitle: (r.item as T).Title ?? "Unknown",
      user,
      previousValue: fieldChanges?.previousValue,
      newValue: fieldChanges?.newValue,
      status: r.status === ItemOperationStatus.Succeeded
        ? "Success" as const
        : r.status === ItemOperationStatus.Cancelled
          ? "Cancelled" as const
          : "Failed" as const,
      error: r.error,
      batchId,
    }));

    this._auditLog.push(...entries);
    return entries;
  }

  /**
   * Persist audit entries to the SharePoint Workflow Audit Log list.
   */
  public async persistAuditLog(
    listTitle: string = "Workflow Audit Log",
    entries?: AuditEntry[]
  ): Promise<void> {
    const toWrite = entries ?? this._auditLog;
    if (toWrite.length === 0) return;

    const [batchedSP, execute] = this._sp.batched();

    for (const entry of toWrite) {
      batchedSP.web.lists.getByTitle(listTitle).items.add({
        Title: `${entry.operation} - ${entry.itemTitle}`,
        Operation: entry.operation,
        ItemID: String(entry.itemId),
        User: entry.user,
        PreviousValue: entry.previousValue ?? "",
        NewValue: entry.newValue ?? "",
        Status: entry.status,
        Error: entry.error ?? "",
        BatchId: entry.batchId,
        Timestamp: entry.timestamp.toISOString(),
      });
    }

    await execute();

    // Clear persisted entries from internal log
    if (!entries) {
      this._auditLog.length = 0;
    }
  }

  /**
   * Get the accumulated audit log entries (not yet persisted).
   */
  public getAuditLog(): ReadonlyArray<AuditEntry> {
    return [...this._auditLog];
  }

  /**
   * Clear the internal audit log.
   */
  public clearAuditLog(): void {
    this._auditLog.length = 0;
  }

  // -----------------------------------------------------------------------
  // Private helpers
  // -----------------------------------------------------------------------

  /** Split an array into chunks of the given size. */
  private _chunk<U>(arr: U[], size: number): U[][] {
    const chunks: U[][] = [];
    for (let i = 0; i < arr.length; i += size) {
      chunks.push(arr.slice(i, i + size));
    }
    return chunks;
  }

  /** Calculate exponential backoff delay: baseDelay * 2^(attempt-1) with jitter. */
  private _calculateBackoff(attempt: number, baseDelay: number): number {
    const exponential = baseDelay * Math.pow(2, attempt - 1);
    const jitter = Math.random() * baseDelay * 0.5;
    return Math.min(exponential + jitter, 30000); // Cap at 30 seconds
  }

  /** Wait for the specified duration, abortable via cancellation token. */
  private _delay(ms: number, cancellationToken?: AbortSignal): Promise<void> {
    return new Promise((resolve) => {
      if (cancellationToken?.aborted) {
        resolve();
        return;
      }

      const timer = setTimeout(resolve, ms);

      cancellationToken?.addEventListener("abort", () => {
        clearTimeout(timer);
        resolve();
      }, { once: true });
    });
  }

  /** Check if an error is a SharePoint 429 throttling response. */
  private _isThrottled(err: unknown): boolean {
    if (err && typeof err === "object") {
      const httpError = err as { status?: number; response?: { status?: number } };
      return httpError.status === 429 || httpError.response?.status === 429;
    }
    return false;
  }

  /** Extract the Retry-After header value in milliseconds. Default: 5000ms. */
  private _getRetryAfterMs(err: unknown): number {
    const DEFAULT_RETRY_MS = 5000;
    if (err && typeof err === "object") {
      const httpError = err as { response?: { headers?: { get?: (k: string) => string | null } } };
      const retryAfter = httpError.response?.headers?.get?.("Retry-After");
      if (retryAfter) {
        const seconds = parseInt(retryAfter, 10);
        if (!isNaN(seconds)) {
          return seconds * 1000;
        }
      }
    }
    return DEFAULT_RETRY_MS;
  }

  /** Extract a human-readable error message from an unknown error. */
  private _extractErrorMessage(err: unknown): string {
    if (err instanceof Error) return err.message;
    if (typeof err === "string") return err;
    if (err && typeof err === "object") {
      const errObj = err as { message?: string; data?: { "odata.error"?: { message?: { value?: string } } } };
      if (errObj.data?.["odata.error"]?.message?.value) {
        return errObj.data["odata.error"].message.value;
      }
      if (errObj.message) return errObj.message;
    }
    return "An unknown error occurred";
  }

  /** Generate a unique batch operation ID. */
  private _generateBatchId(): string {
    return `batch-${Date.now().toString(36)}-${Math.random().toString(36).substring(2, 8)}`;
  }
}
