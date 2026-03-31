// ---------------------------------------------------------------------------
// Status Transition Service -- Type-safe state machine for item status
// ---------------------------------------------------------------------------

/**
 * All valid status values in the workflow system.
 * Used as discriminated union keys throughout the state machine.
 */
export type WorkflowStatus =
  | "Not Started"
  | "In Progress"
  | "Under Review"
  | "Approved"
  | "Rejected"
  | "On Hold"
  | "Cancelled"
  | "Archived";

/**
 * Side effect action triggered when a status transition occurs.
 */
export interface TransitionSideEffect {
  /** Unique identifier for the side effect. */
  id: string;
  /** Human-readable description. */
  description: string;
  /** The async function to execute. Receives transition context. */
  execute: (context: TransitionContext) => Promise<void>;
}

/**
 * Context passed to side effects and validators during a transition.
 */
export interface TransitionContext {
  /** The SharePoint list item ID. */
  itemId: number;
  /** The item title for logging/display. */
  itemTitle: string;
  /** The list title where the item resides. */
  listTitle: string;
  /** The status being transitioned from. */
  fromStatus: WorkflowStatus;
  /** The status being transitioned to. */
  toStatus: WorkflowStatus;
  /** The user performing the transition. */
  performedBy: string;
  /** Timestamp of the transition. */
  timestamp: Date;
  /** Optional reason/comment for the transition. */
  reason?: string;
  /** Additional metadata. */
  metadata?: Record<string, unknown>;
}

/**
 * Definition of a single allowed transition between statuses.
 */
export interface TransitionDefinition {
  /** The source status. */
  from: WorkflowStatus;
  /** The target status. */
  to: WorkflowStatus;
  /** Optional guard condition. Return false to block the transition. */
  guard?: (context: TransitionContext) => Promise<boolean>;
  /** Optional human-readable label for the guard. */
  guardDescription?: string;
  /** Side effects to execute after the transition succeeds. */
  sideEffects: TransitionSideEffect[];
  /** Whether a reason/comment is required for this transition. */
  requiresReason: boolean;
  /** Roles allowed to perform this transition (empty = all). */
  allowedRoles: string[];
}

/**
 * Result of a transition validation.
 */
export interface TransitionValidationResult {
  /** Whether the transition is allowed. */
  isValid: boolean;
  /** Error messages if not valid. */
  errors: string[];
  /** Warning messages (transition allowed but with caveats). */
  warnings: string[];
}

/**
 * Result of executing a transition.
 */
export interface TransitionResult {
  /** Whether the transition completed successfully. */
  success: boolean;
  /** The transition context. */
  context: TransitionContext;
  /** Side effects that executed successfully. */
  completedSideEffects: string[];
  /** Side effects that failed (transition still recorded). */
  failedSideEffects: { id: string; error: string }[];
  /** Error message if the transition itself failed. */
  error?: string;
}

/**
 * An entry in the status transition audit trail.
 */
export interface TransitionAuditEntry {
  id: string;
  timestamp: Date;
  itemId: number;
  itemTitle: string;
  listTitle: string;
  fromStatus: WorkflowStatus;
  toStatus: WorkflowStatus;
  performedBy: string;
  reason?: string;
  sideEffectsExecuted: string[];
  sideEffectsFailed: string[];
  success: boolean;
}

// ---------------------------------------------------------------------------
// Default Transition Map
// ---------------------------------------------------------------------------

/**
 * Default transition rules for the workflow system.
 * These can be overridden or extended via `registerTransition()`.
 *
 * Valid transitions:
 *
 *   Not Started  --> In Progress
 *   Not Started  --> Cancelled
 *   In Progress  --> Under Review
 *   In Progress  --> On Hold
 *   In Progress  --> Cancelled
 *   Under Review --> Approved
 *   Under Review --> Rejected
 *   Under Review --> In Progress  (sent back for revision)
 *   Approved     --> Archived
 *   Rejected     --> In Progress  (resubmitted)
 *   Rejected     --> Cancelled
 *   On Hold      --> In Progress  (resumed)
 *   On Hold      --> Cancelled
 */
const DEFAULT_TRANSITIONS: Omit<TransitionDefinition, "sideEffects" | "guard">[] = [
  { from: "Not Started", to: "In Progress", requiresReason: false, allowedRoles: [] },
  { from: "Not Started", to: "Cancelled", requiresReason: true, allowedRoles: [] },
  { from: "In Progress", to: "Under Review", requiresReason: false, allowedRoles: [] },
  { from: "In Progress", to: "On Hold", requiresReason: true, allowedRoles: [] },
  { from: "In Progress", to: "Cancelled", requiresReason: true, allowedRoles: [] },
  { from: "Under Review", to: "Approved", requiresReason: false, allowedRoles: ["Approver", "Admin"] },
  { from: "Under Review", to: "Rejected", requiresReason: true, allowedRoles: ["Approver", "Admin"] },
  { from: "Under Review", to: "In Progress", requiresReason: true, allowedRoles: ["Approver", "Admin"] },
  { from: "Approved", to: "Archived", requiresReason: false, allowedRoles: ["Admin"] },
  { from: "Rejected", to: "In Progress", requiresReason: false, allowedRoles: [] },
  { from: "Rejected", to: "Cancelled", requiresReason: true, allowedRoles: [] },
  { from: "On Hold", to: "In Progress", requiresReason: false, allowedRoles: [] },
  { from: "On Hold", to: "Cancelled", requiresReason: true, allowedRoles: [] },
];

// ---------------------------------------------------------------------------
// Service
// ---------------------------------------------------------------------------

/**
 * State machine for SharePoint list item status transitions.
 *
 * Features:
 * - Defines and validates legal status transitions
 * - Executes side effects on successful transitions (notifications, field updates)
 * - Maintains a full audit trail of all transitions
 * - Supports custom guard conditions and role-based access
 * - Type-safe with TypeScript discriminated unions
 *
 * @example
 * ```typescript
 * const service = new StatusTransitionService();
 *
 * // Register a side effect for approval
 * service.registerSideEffect("Under Review", "Approved", {
 *   id: "send-approval-notification",
 *   description: "Send approval email to submitter",
 *   execute: async (ctx) => {
 *     await sendEmail(ctx.performedBy, `Item ${ctx.itemTitle} approved`);
 *   },
 * });
 *
 * // Validate before performing
 * const validation = await service.validate("Under Review", "Approved", context);
 * if (validation.isValid) {
 *   const result = await service.transition(context);
 * }
 * ```
 */
export class StatusTransitionService {
  private readonly _transitions: Map<string, TransitionDefinition> = new Map();
  private readonly _auditTrail: TransitionAuditEntry[] = [];

  constructor(useDefaults: boolean = true) {
    if (useDefaults) {
      this._registerDefaults();
    }
  }

  // -----------------------------------------------------------------------
  // Configuration API
  // -----------------------------------------------------------------------

  /**
   * Register a custom transition definition. Overwrites any existing
   * transition for the same from/to pair.
   */
  public registerTransition(definition: TransitionDefinition): void {
    const key = this._transitionKey(definition.from, definition.to);
    this._transitions.set(key, definition);
  }

  /**
   * Remove a registered transition.
   */
  public removeTransition(from: WorkflowStatus, to: WorkflowStatus): boolean {
    return this._transitions.delete(this._transitionKey(from, to));
  }

  /**
   * Register a side effect for an existing transition.
   * Throws if the transition does not exist.
   */
  public registerSideEffect(
    from: WorkflowStatus,
    to: WorkflowStatus,
    sideEffect: TransitionSideEffect
  ): void {
    const key = this._transitionKey(from, to);
    const transition = this._transitions.get(key);
    if (!transition) {
      throw new Error(
        `Cannot register side effect: no transition defined from "${from}" to "${to}".`
      );
    }
    // Prevent duplicate side effect IDs
    if (transition.sideEffects.some((se) => se.id === sideEffect.id)) {
      throw new Error(
        `Side effect with id "${sideEffect.id}" already registered for ${from} -> ${to}.`
      );
    }
    transition.sideEffects.push(sideEffect);
  }

  /**
   * Set a guard condition on an existing transition.
   */
  public setGuard(
    from: WorkflowStatus,
    to: WorkflowStatus,
    guard: (context: TransitionContext) => Promise<boolean>,
    description?: string
  ): void {
    const key = this._transitionKey(from, to);
    const transition = this._transitions.get(key);
    if (!transition) {
      throw new Error(
        `Cannot set guard: no transition defined from "${from}" to "${to}".`
      );
    }
    transition.guard = guard;
    if (description) {
      transition.guardDescription = description;
    }
  }

  // -----------------------------------------------------------------------
  // Query API
  // -----------------------------------------------------------------------

  /**
   * Get all valid target statuses from the given current status.
   */
  public getAvailableTransitions(fromStatus: WorkflowStatus): WorkflowStatus[] {
    const available: WorkflowStatus[] = [];
    for (const [, def] of this._transitions) {
      if (def.from === fromStatus) {
        available.push(def.to);
      }
    }
    return available;
  }

  /**
   * Get the full transition definition for a from/to pair, or undefined if
   * the transition is not registered.
   */
  public getTransitionDefinition(
    from: WorkflowStatus,
    to: WorkflowStatus
  ): TransitionDefinition | undefined {
    return this._transitions.get(this._transitionKey(from, to));
  }

  /**
   * Check whether a transition is structurally defined (ignoring guards/roles).
   */
  public isTransitionDefined(from: WorkflowStatus, to: WorkflowStatus): boolean {
    return this._transitions.has(this._transitionKey(from, to));
  }

  /**
   * Get all registered transitions as an array.
   */
  public getAllTransitions(): TransitionDefinition[] {
    return Array.from(this._transitions.values());
  }

  // -----------------------------------------------------------------------
  // Validation
  // -----------------------------------------------------------------------

  /**
   * Validate whether a transition is allowed, checking:
   * 1. The transition is defined in the state machine.
   * 2. A reason is provided if required.
   * 3. The user has the required role.
   * 4. The guard condition passes (if defined).
   */
  public async validate(
    from: WorkflowStatus,
    to: WorkflowStatus,
    context: TransitionContext,
    userRoles?: string[]
  ): Promise<TransitionValidationResult> {
    const errors: string[] = [];
    const warnings: string[] = [];

    // 1. Check if transition is defined
    const definition = this.getTransitionDefinition(from, to);
    if (!definition) {
      errors.push(
        `Transition from "${from}" to "${to}" is not allowed. ` +
        `Valid transitions from "${from}": ${this.getAvailableTransitions(from).join(", ") || "none"}.`
      );
      return { isValid: false, errors, warnings };
    }

    // 2. Check reason requirement
    if (definition.requiresReason && (!context.reason || context.reason.trim().length === 0)) {
      errors.push(`A reason is required for transitioning from "${from}" to "${to}".`);
    }

    // 3. Check role-based access
    if (definition.allowedRoles.length > 0 && userRoles) {
      const hasRole = definition.allowedRoles.some((role) => userRoles.includes(role));
      if (!hasRole) {
        errors.push(
          `User does not have the required role. ` +
          `Allowed roles: ${definition.allowedRoles.join(", ")}.`
        );
      }
    } else if (definition.allowedRoles.length > 0 && !userRoles) {
      warnings.push("Role validation skipped: no user roles provided.");
    }

    // 4. Check guard condition
    if (definition.guard && errors.length === 0) {
      try {
        const guardResult = await definition.guard(context);
        if (!guardResult) {
          errors.push(
            definition.guardDescription
              ? `Guard condition failed: ${definition.guardDescription}`
              : `Guard condition blocked the transition from "${from}" to "${to}".`
          );
        }
      } catch (err) {
        errors.push(
          `Guard condition threw an error: ${err instanceof Error ? err.message : String(err)}`
        );
      }
    }

    return {
      isValid: errors.length === 0,
      errors,
      warnings,
    };
  }

  // -----------------------------------------------------------------------
  // Execution
  // -----------------------------------------------------------------------

  /**
   * Execute a status transition: validate, record, and run side effects.
   *
   * This method does NOT update the SharePoint item directly. It validates
   * the transition, records it in the audit trail, and executes side effects.
   * The caller is responsible for the actual field update.
   *
   * @returns TransitionResult indicating success/failure and side effect outcomes.
   */
  public async transition(
    context: TransitionContext,
    userRoles?: string[]
  ): Promise<TransitionResult> {
    // Validate first
    const validation = await this.validate(
      context.fromStatus,
      context.toStatus,
      context,
      userRoles
    );

    if (!validation.isValid) {
      return {
        success: false,
        context,
        completedSideEffects: [],
        failedSideEffects: [],
        error: validation.errors.join(" | "),
      };
    }

    const definition = this.getTransitionDefinition(context.fromStatus, context.toStatus)!;
    const completedSideEffects: string[] = [];
    const failedSideEffects: { id: string; error: string }[] = [];

    // Execute side effects (non-blocking: a failed side effect does not
    // prevent the transition from being recorded)
    for (const sideEffect of definition.sideEffects) {
      try {
        await sideEffect.execute(context);
        completedSideEffects.push(sideEffect.id);
      } catch (err) {
        failedSideEffects.push({
          id: sideEffect.id,
          error: err instanceof Error ? err.message : String(err),
        });
      }
    }

    // Record in audit trail
    const auditEntry: TransitionAuditEntry = {
      id: `txn-${Date.now().toString(36)}-${Math.random().toString(36).substring(2, 6)}`,
      timestamp: context.timestamp,
      itemId: context.itemId,
      itemTitle: context.itemTitle,
      listTitle: context.listTitle,
      fromStatus: context.fromStatus,
      toStatus: context.toStatus,
      performedBy: context.performedBy,
      reason: context.reason,
      sideEffectsExecuted: completedSideEffects,
      sideEffectsFailed: failedSideEffects.map((f) => f.id),
      success: true,
    };

    this._auditTrail.push(auditEntry);

    return {
      success: true,
      context,
      completedSideEffects,
      failedSideEffects,
    };
  }

  // -----------------------------------------------------------------------
  // Audit Trail
  // -----------------------------------------------------------------------

  /**
   * Get the full audit trail.
   */
  public getAuditTrail(): ReadonlyArray<TransitionAuditEntry> {
    return [...this._auditTrail];
  }

  /**
   * Get audit trail entries for a specific item.
   */
  public getItemHistory(itemId: number): TransitionAuditEntry[] {
    return this._auditTrail.filter((e) => e.itemId === itemId);
  }

  /**
   * Clear the audit trail (e.g., after persisting to SharePoint).
   */
  public clearAuditTrail(): void {
    this._auditTrail.length = 0;
  }

  // -----------------------------------------------------------------------
  // Private helpers
  // -----------------------------------------------------------------------

  /** Create a unique key for a from/to transition pair. */
  private _transitionKey(from: WorkflowStatus, to: WorkflowStatus): string {
    return `${from}|${to}`;
  }

  /** Register the default transition definitions. */
  private _registerDefaults(): void {
    for (const def of DEFAULT_TRANSITIONS) {
      this.registerTransition({
        ...def,
        sideEffects: [],
      });
    }
  }
}
