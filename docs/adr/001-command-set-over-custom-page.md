# ADR 001: Use SPFx Command Set Over Custom Application Page for Bulk Actions

**Status:** Accepted
**Date:** 2026-01-15
**Decision Makers:** Solution Architect, SharePoint Lead

## Context

The solution requires bulk operations (approve, export, assign) on SharePoint list items. Two primary approaches exist:

1. **SPFx ListView Command Set** -- Toolbar buttons injected directly into the SharePoint list view.
2. **SPFx Application Customizer / Custom Page** -- A standalone page (or top/bottom placeholder) with its own UI for managing items.

We need to decide which pattern best balances usability, maintainability, and native SharePoint integration.

## Decision

We will use an **SPFx ListView Command Set** extension to provide bulk action capabilities.

## Rationale

### Pros of Command Set

- **Native integration**: Buttons appear in the standard SharePoint list command bar alongside built-in actions (New, Edit, Share). Users do not navigate away from the list.
- **Context-aware**: The extension receives `event.selectedRows` automatically, providing the selected items' field values with zero additional API calls.
- **No separate deployment surface**: The extension is deployed as part of the `.sppkg` package and automatically associated with the target lists via `ClientSideComponentId` in the PnP provisioning template. No custom pages or navigation nodes to manage.
- **Multi-list reuse**: A single manifest with `listType` scoping allows the same extension to serve Project Tracker, Document Approval Queue, and Change Request Log without duplication.
- **Familiar UX**: Users see toolbar buttons exactly where they expect them. No training required beyond "select items, click button."
- **Lazy loading**: SPFx Command Set JS bundles load only when the list view renders, avoiding upfront cost on site navigation.

### Cons / Risks

- **Limited layout control**: The command bar provides only buttons (with optional dropdowns). Complex forms require a secondary panel (which we use for Assign To and Export Dialog).
- **Dependency on list view context**: The extension only works within a list view web part or the full-page list experience. It cannot operate from a dashboard page or Teams tab without embedding the list.
- **Testing complexity**: SPFx extensions require either `gulp serve` with `serveConfigurations` or deployment to a test site. Unit testing the command handler logic requires mocking `BaseListViewCommandSet` context.

### Why Not a Custom Application Page

- A custom page requires its own routing, navigation entry, and separate data-fetching logic (querying the list from scratch rather than leveraging the already-loaded list view).
- Users must leave their current context to perform bulk actions, breaking the "stay in the list" workflow.
- Additional deployment surface increases maintenance burden and potential for configuration drift.
- The custom page pattern is better suited for scenarios where the UI is fundamentally different from a list (e.g., a Kanban board or calendar view), which is not the case here.

## Consequences

- All bulk actions are initiated from the list view toolbar.
- Complex interactions (people picker, export column selection) open in Fluent UI side panels overlaid on the list view.
- The extension must handle `onListViewUpdated` to enable/disable buttons based on selection count and item status.
- PnP provisioning templates include `CustomAction` entries to associate the Command Set with each target list.
- Future dashboard requirements (if any) would be implemented as a separate SPFx web part, not by changing the bulk action pattern.

## Alternatives Considered

| Alternative | Reason for Rejection |
|---|---|
| Custom Application Page | Breaks "stay in list" UX; extra deployment surface |
| Power Apps embedded form | Cannot access list selection context; separate licensing |
| Power Automate button column | One item at a time; no multi-select |
| SharePoint JSON column formatting + Flow | Limited to simple actions; no complex UI |
