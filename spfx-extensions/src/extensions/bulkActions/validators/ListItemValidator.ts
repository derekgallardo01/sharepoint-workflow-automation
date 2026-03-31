// ---------------------------------------------------------------------------
// List Item Validation Framework
// ---------------------------------------------------------------------------

/**
 * Severity level for a validation issue.
 */
export enum ValidationSeverity {
  Error = "Error",
  Warning = "Warning",
}

/**
 * A single validation issue on a specific field.
 */
export interface ValidationIssue {
  /** The internal name of the field. */
  field: string;
  /** The display name of the field (for user-facing messages). */
  fieldDisplayName: string;
  /** Human-readable error/warning message. */
  message: string;
  /** Severity level. */
  severity: ValidationSeverity;
  /** The validator rule that produced this issue. */
  rule: string;
  /** The invalid value (for debugging). */
  actualValue?: unknown;
}

/**
 * Result of validating a single item.
 */
export interface ItemValidationResult {
  /** Whether the item passed all Error-level validations. */
  isValid: boolean;
  /** The item that was validated. */
  item: Record<string, unknown>;
  /** All issues found (errors and warnings). */
  issues: ValidationIssue[];
  /** Shortcut: only Error-level issues. */
  errors: ValidationIssue[];
  /** Shortcut: only Warning-level issues. */
  warnings: ValidationIssue[];
}

/**
 * Result of validating a batch of items.
 */
export interface BatchValidationResult {
  /** Whether all items passed validation. */
  allValid: boolean;
  /** Total items validated. */
  total: number;
  /** Number of valid items. */
  validCount: number;
  /** Number of invalid items (at least one Error). */
  invalidCount: number;
  /** Per-item validation results. */
  results: ItemValidationResult[];
  /** Items that passed validation (convenience accessor). */
  validItems: Record<string, unknown>[];
  /** Items that failed validation (convenience accessor). */
  invalidItems: Record<string, unknown>[];
}

// ---------------------------------------------------------------------------
// Validation Rule Types
// ---------------------------------------------------------------------------

/**
 * Base interface for a validation rule.
 */
export interface ValidationRule {
  /** Unique rule identifier. */
  id: string;
  /** The field internal name this rule applies to. */
  field: string;
  /** The field display name for error messages. */
  fieldDisplayName: string;
  /** Severity: Error blocks the operation, Warning is informational. */
  severity: ValidationSeverity;
  /** The validation function. Returns null if valid, or an error message string. */
  validate: (value: unknown, item: Record<string, unknown>) => string | null;
}

/**
 * Schema definition for validating items of a particular content type.
 * Maps content type name to an array of validation rules.
 */
export interface ValidationSchema {
  /** Name of the content type or list this schema applies to. */
  contentType: string;
  /** Display name for error reporting. */
  displayName: string;
  /** Validation rules to apply. */
  rules: ValidationRule[];
}

// ---------------------------------------------------------------------------
// Built-in Validator Factories
// ---------------------------------------------------------------------------

/**
 * Built-in validator functions. Each factory returns a `ValidationRule`.
 */
export const Validators = {
  /**
   * Validates that a field has a non-empty value.
   */
  required(
    field: string,
    displayName: string,
    severity: ValidationSeverity = ValidationSeverity.Error
  ): ValidationRule {
    return {
      id: `required:${field}`,
      field,
      fieldDisplayName: displayName,
      severity,
      validate: (value: unknown): string | null => {
        if (value === null || value === undefined) {
          return `${displayName} is required.`;
        }
        if (typeof value === "string" && value.trim().length === 0) {
          return `${displayName} is required.`;
        }
        return null;
      },
    };
  },

  /**
   * Validates that a string field does not exceed a maximum length.
   */
  maxLength(
    field: string,
    displayName: string,
    max: number,
    severity: ValidationSeverity = ValidationSeverity.Error
  ): ValidationRule {
    return {
      id: `maxLength:${field}:${max}`,
      field,
      fieldDisplayName: displayName,
      severity,
      validate: (value: unknown): string | null => {
        if (value === null || value === undefined) return null; // Not this validator's concern
        const str = String(value);
        if (str.length > max) {
          return `${displayName} must be at most ${max} characters (currently ${str.length}).`;
        }
        return null;
      },
    };
  },

  /**
   * Validates that a string field has at least a minimum length.
   */
  minLength(
    field: string,
    displayName: string,
    min: number,
    severity: ValidationSeverity = ValidationSeverity.Error
  ): ValidationRule {
    return {
      id: `minLength:${field}:${min}`,
      field,
      fieldDisplayName: displayName,
      severity,
      validate: (value: unknown): string | null => {
        if (value === null || value === undefined) return null;
        const str = String(value);
        if (str.length < min) {
          return `${displayName} must be at least ${min} characters.`;
        }
        return null;
      },
    };
  },

  /**
   * Validates that a date field falls within a specified range.
   */
  dateRange(
    field: string,
    displayName: string,
    options: { min?: Date; max?: Date },
    severity: ValidationSeverity = ValidationSeverity.Error
  ): ValidationRule {
    return {
      id: `dateRange:${field}`,
      field,
      fieldDisplayName: displayName,
      severity,
      validate: (value: unknown): string | null => {
        if (value === null || value === undefined) return null;

        const date = value instanceof Date ? value : new Date(String(value));
        if (isNaN(date.getTime())) {
          return `${displayName} contains an invalid date.`;
        }

        if (options.min && date < options.min) {
          return `${displayName} must be on or after ${options.min.toLocaleDateString()}.`;
        }
        if (options.max && date > options.max) {
          return `${displayName} must be on or before ${options.max.toLocaleDateString()}.`;
        }
        return null;
      },
    };
  },

  /**
   * Validates that a field value is one of the allowed choices.
   */
  choiceValues(
    field: string,
    displayName: string,
    allowedValues: string[],
    severity: ValidationSeverity = ValidationSeverity.Error
  ): ValidationRule {
    const lowerAllowed = allowedValues.map((v) => v.toLowerCase());
    return {
      id: `choiceValues:${field}`,
      field,
      fieldDisplayName: displayName,
      severity,
      validate: (value: unknown): string | null => {
        if (value === null || value === undefined) return null;
        const str = String(value).toLowerCase();
        if (!lowerAllowed.includes(str)) {
          return `${displayName} must be one of: ${allowedValues.join(", ")}. Got: "${value}".`;
        }
        return null;
      },
    };
  },

  /**
   * Validates that a Person field references an existing user.
   * Checks that the value is a non-zero user ID or a non-empty object with Id/EMail.
   */
  personExists(
    field: string,
    displayName: string,
    severity: ValidationSeverity = ValidationSeverity.Error
  ): ValidationRule {
    return {
      id: `personExists:${field}`,
      field,
      fieldDisplayName: displayName,
      severity,
      validate: (value: unknown): string | null => {
        if (value === null || value === undefined) return null;

        // SharePoint person fields can be numeric (ID) or object { Id, Title, EMail }
        if (typeof value === "number") {
          if (value <= 0) {
            return `${displayName} must reference a valid user.`;
          }
          return null;
        }

        if (typeof value === "object" && value !== null) {
          const person = value as { Id?: number; EMail?: string };
          if (person.Id && person.Id > 0) return null;
          if (person.EMail && person.EMail.length > 0) return null;
          return `${displayName} must reference a valid user.`;
        }

        return `${displayName} has an unrecognized format.`;
      },
    };
  },

  /**
   * Validates that a numeric field falls within a specified range.
   */
  numberRange(
    field: string,
    displayName: string,
    options: { min?: number; max?: number },
    severity: ValidationSeverity = ValidationSeverity.Error
  ): ValidationRule {
    return {
      id: `numberRange:${field}`,
      field,
      fieldDisplayName: displayName,
      severity,
      validate: (value: unknown): string | null => {
        if (value === null || value === undefined) return null;
        const num = typeof value === "number" ? value : parseFloat(String(value));
        if (isNaN(num)) {
          return `${displayName} must be a valid number.`;
        }
        if (options.min !== undefined && num < options.min) {
          return `${displayName} must be at least ${options.min}.`;
        }
        if (options.max !== undefined && num > options.max) {
          return `${displayName} must be at most ${options.max}.`;
        }
        return null;
      },
    };
  },

  /**
   * Validates that a field matches a regular expression pattern.
   */
  pattern(
    field: string,
    displayName: string,
    regex: RegExp,
    errorMessage: string,
    severity: ValidationSeverity = ValidationSeverity.Error
  ): ValidationRule {
    return {
      id: `pattern:${field}`,
      field,
      fieldDisplayName: displayName,
      severity,
      validate: (value: unknown): string | null => {
        if (value === null || value === undefined) return null;
        const str = String(value);
        if (!regex.test(str)) {
          return errorMessage || `${displayName} does not match the required pattern.`;
        }
        return null;
      },
    };
  },

  /**
   * Creates a custom validation rule with an arbitrary validation function.
   */
  custom(
    field: string,
    displayName: string,
    ruleId: string,
    validateFn: (value: unknown, item: Record<string, unknown>) => string | null,
    severity: ValidationSeverity = ValidationSeverity.Error
  ): ValidationRule {
    return {
      id: `custom:${ruleId}`,
      field,
      fieldDisplayName: displayName,
      severity,
      validate: validateFn,
    };
  },
};

// ---------------------------------------------------------------------------
// Validator Service
// ---------------------------------------------------------------------------

/**
 * Validation framework for SharePoint list items.
 *
 * Features:
 * - Schema-based validation with rules defined per content type
 * - Built-in validators: required, maxLength, minLength, dateRange,
 *   choiceValues, personExists, numberRange, pattern
 * - Custom validator support via `Validators.custom()`
 * - Returns typed `ValidationResult` with field-level errors and warnings
 * - Batch validation for multiple items before bulk operations
 *
 * @example
 * ```typescript
 * const validator = new ListItemValidator();
 *
 * // Define a schema for the Project Tracker content type
 * validator.registerSchema({
 *   contentType: "ProjectItem",
 *   displayName: "Project Tracker",
 *   rules: [
 *     Validators.required("Title", "Project Title"),
 *     Validators.maxLength("Title", "Project Title", 255),
 *     Validators.choiceValues("Status", "Status", [
 *       "Not Started", "In Progress", "Under Review", "Approved", "Rejected"
 *     ]),
 *     Validators.required("AssignedToId", "Assigned To"),
 *     Validators.personExists("AssignedToId", "Assigned To"),
 *     Validators.dateRange("DueDate", "Due Date", { min: new Date() }),
 *     Validators.numberRange("PercentComplete", "% Complete", { min: 0, max: 100 }),
 *   ],
 * });
 *
 * // Validate a batch of items before bulk approve
 * const result = validator.validateBatch(selectedItems, "ProjectItem");
 * if (!result.allValid) {
 *   // Show validation errors in the ProgressPanel
 *   result.invalidItems.forEach(item => { ... });
 * }
 * ```
 */
export class ListItemValidator {
  private readonly _schemas: Map<string, ValidationSchema> = new Map();

  // -----------------------------------------------------------------------
  // Schema Management
  // -----------------------------------------------------------------------

  /**
   * Register a validation schema for a content type.
   * Overwrites any previously registered schema for the same content type.
   */
  public registerSchema(schema: ValidationSchema): void {
    this._schemas.set(schema.contentType, schema);
  }

  /**
   * Get a registered schema by content type name.
   */
  public getSchema(contentType: string): ValidationSchema | undefined {
    return this._schemas.get(contentType);
  }

  /**
   * Remove a registered schema.
   */
  public removeSchema(contentType: string): boolean {
    return this._schemas.delete(contentType);
  }

  /**
   * Get all registered schema names.
   */
  public getRegisteredContentTypes(): string[] {
    return Array.from(this._schemas.keys());
  }

  /**
   * Add a rule to an existing schema. Creates the schema if it does not exist.
   */
  public addRule(contentType: string, rule: ValidationRule): void {
    let schema = this._schemas.get(contentType);
    if (!schema) {
      schema = { contentType, displayName: contentType, rules: [] };
      this._schemas.set(contentType, schema);
    }
    // Prevent duplicate rule IDs
    if (schema.rules.some((r) => r.id === rule.id)) {
      throw new Error(
        `Rule "${rule.id}" already exists in schema "${contentType}". Remove it first.`
      );
    }
    schema.rules.push(rule);
  }

  /**
   * Remove a rule from a schema by rule ID.
   */
  public removeRule(contentType: string, ruleId: string): boolean {
    const schema = this._schemas.get(contentType);
    if (!schema) return false;
    const idx = schema.rules.findIndex((r) => r.id === ruleId);
    if (idx === -1) return false;
    schema.rules.splice(idx, 1);
    return true;
  }

  // -----------------------------------------------------------------------
  // Validation
  // -----------------------------------------------------------------------

  /**
   * Validate a single item against the specified content type schema.
   *
   * @param item        - The SharePoint list item (field name -> value).
   * @param contentType - The content type schema to validate against.
   * @returns Typed validation result with field-level issues.
   */
  public validateItem(
    item: Record<string, unknown>,
    contentType: string
  ): ItemValidationResult {
    const schema = this._schemas.get(contentType);
    if (!schema) {
      throw new Error(
        `No validation schema registered for content type "${contentType}". ` +
        `Registered types: ${this.getRegisteredContentTypes().join(", ") || "none"}.`
      );
    }

    const issues: ValidationIssue[] = [];

    for (const rule of schema.rules) {
      const value = item[rule.field];
      const errorMessage = rule.validate(value, item);

      if (errorMessage !== null) {
        issues.push({
          field: rule.field,
          fieldDisplayName: rule.fieldDisplayName,
          message: errorMessage,
          severity: rule.severity,
          rule: rule.id,
          actualValue: value,
        });
      }
    }

    const errors = issues.filter((i) => i.severity === ValidationSeverity.Error);
    const warnings = issues.filter((i) => i.severity === ValidationSeverity.Warning);

    return {
      isValid: errors.length === 0,
      item,
      issues,
      errors,
      warnings,
    };
  }

  /**
   * Validate a batch of items, returning per-item results and aggregate counts.
   * Used before bulk operations to prevent invalid data from being submitted.
   *
   * @param items       - Array of SharePoint list items.
   * @param contentType - The content type schema to validate against.
   * @returns Batch validation result with valid/invalid item lists.
   */
  public validateBatch(
    items: Record<string, unknown>[],
    contentType: string
  ): BatchValidationResult {
    const results: ItemValidationResult[] = items.map((item) =>
      this.validateItem(item, contentType)
    );

    const validItems = results.filter((r) => r.isValid).map((r) => r.item);
    const invalidItems = results.filter((r) => !r.isValid).map((r) => r.item);

    return {
      allValid: invalidItems.length === 0,
      total: items.length,
      validCount: validItems.length,
      invalidCount: invalidItems.length,
      results,
      validItems,
      invalidItems,
    };
  }

  /**
   * Validate a single item with ad-hoc rules (no schema registration needed).
   * Useful for one-off validations or dynamic rule sets.
   */
  public validateWithRules(
    item: Record<string, unknown>,
    rules: ValidationRule[]
  ): ItemValidationResult {
    const issues: ValidationIssue[] = [];

    for (const rule of rules) {
      const value = item[rule.field];
      const errorMessage = rule.validate(value, item);

      if (errorMessage !== null) {
        issues.push({
          field: rule.field,
          fieldDisplayName: rule.fieldDisplayName,
          message: errorMessage,
          severity: rule.severity,
          rule: rule.id,
          actualValue: value,
        });
      }
    }

    const errors = issues.filter((i) => i.severity === ValidationSeverity.Error);
    const warnings = issues.filter((i) => i.severity === ValidationSeverity.Warning);

    return {
      isValid: errors.length === 0,
      item,
      issues,
      errors,
      warnings,
    };
  }
}
