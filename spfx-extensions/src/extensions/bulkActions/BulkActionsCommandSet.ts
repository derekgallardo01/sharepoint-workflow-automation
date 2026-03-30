import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseListViewCommandSet,
  Command,
  IListViewCommandSetExecuteEventParameters,
  ListViewStateChangedEventArgs,
  RowAccessor
} from '@microsoft/sp-listview-extensibility';
import { Dialog } from '@microsoft/sp-dialog';
import { spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/batching';

const LOG_SOURCE: string = 'BulkActionsCommandSet';

export interface IBulkActionsCommandSetProperties {
  statusFieldName: string;
}

export default class BulkActionsCommandSet extends BaseListViewCommandSet<IBulkActionsCommandSetProperties> {
  private _sp: ReturnType<typeof spfi>;

  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized BulkActionsCommandSet');

    this._sp = spfi().using(SPFx(this.context));

    const bulkApproveCommand: Command = this.tryGetCommand('BULK_APPROVE');
    const exportCsvCommand: Command = this.tryGetCommand('EXPORT_CSV');
    const assignToCommand: Command = this.tryGetCommand('ASSIGN_TO');

    if (bulkApproveCommand) {
      bulkApproveCommand.visible = false;
    }
    if (exportCsvCommand) {
      exportCsvCommand.visible = false;
    }
    if (assignToCommand) {
      assignToCommand.visible = false;
    }

    this.context.listView.listViewStateChangedEvent.add(
      this,
      this._onListViewStateChanged
    );

    return Promise.resolve();
  }

  private _onListViewStateChanged = (args: ListViewStateChangedEventArgs): void => {
    Log.info(LOG_SOURCE, 'List view state changed');

    const selectedCount = this.context.listView.selectedRows?.length ?? 0;
    const hasSelection = selectedCount > 0;

    const bulkApproveCommand: Command = this.tryGetCommand('BULK_APPROVE');
    if (bulkApproveCommand) {
      bulkApproveCommand.visible = hasSelection;
    }

    const exportCsvCommand: Command = this.tryGetCommand('EXPORT_CSV');
    if (exportCsvCommand) {
      exportCsvCommand.visible = hasSelection;
    }

    const assignToCommand: Command = this.tryGetCommand('ASSIGN_TO');
    if (assignToCommand) {
      assignToCommand.visible = hasSelection;
    }

    this.raiseOnChange();
  }

  @override
  public async onExecute(event: IListViewCommandSetExecuteEventParameters): Promise<void> {
    const selectedRows = this.context.listView.selectedRows ?? [];

    if (selectedRows.length === 0) {
      await Dialog.alert('Please select at least one item.');
      return;
    }

    switch (event.itemId) {
      case 'BULK_APPROVE':
        await this._handleBulkApprove(selectedRows);
        break;
      case 'EXPORT_CSV':
        this._handleExportCsv(selectedRows);
        break;
      case 'ASSIGN_TO':
        await this._handleAssignTo(selectedRows);
        break;
      default:
        throw new Error('Unknown command');
    }
  }

  /**
   * Approves all selected items by updating their Status field to "Approved".
   * Uses PnP JS batching for efficient bulk updates.
   */
  private async _handleBulkApprove(selectedRows: readonly RowAccessor[]): Promise<void> {
    const statusField = this.properties.statusFieldName || 'Status';
    const itemCount = selectedRows.length;

    const confirmed = await this._confirm(
      `Are you sure you want to approve ${itemCount} item(s)?`
    );
    if (!confirmed) {
      return;
    }

    try {
      const listId = this.context.listView.list.guid.toString();
      const [batchedSP, execute] = this._sp.batched();
      const list = batchedSP.web.lists.getById(listId);

      const results: Promise<void>[] = [];

      for (const row of selectedRows) {
        const itemId = parseInt(row.getValueByName('ID'), 10);
        results.push(
          list.items.getById(itemId).update({
            [statusField]: 'Approved'
          }).then(() => undefined)
        );
      }

      await execute();
      await Promise.all(results);

      await Dialog.alert(
        `Successfully approved ${itemCount} item(s). Refresh the page to see updates.`
      );

      Log.info(LOG_SOURCE, `Bulk approved ${itemCount} items`);
    } catch (error) {
      Log.error(LOG_SOURCE, error as Error);
      await Dialog.alert(
        `Error approving items: ${(error as Error).message}. Please try again or contact support.`
      );
    }
  }

  /**
   * Exports selected rows to a CSV file and triggers a browser download.
   * Dynamically reads available columns from the first selected row.
   */
  private _handleExportCsv(selectedRows: readonly RowAccessor[]): void {
    try {
      // Gather field names from the view columns
      const columns = this.context.listView.columns;
      const fieldNames: string[] = [];
      const displayNames: string[] = [];

      for (const column of columns) {
        fieldNames.push(column.field.internalName);
        displayNames.push(column.field.displayName);
      }

      // Build CSV content
      const csvRows: string[] = [];
      csvRows.push(displayNames.map(name => this._escapeCsvField(name)).join(','));

      for (const row of selectedRows) {
        const values: string[] = fieldNames.map(fieldName => {
          const rawValue = row.getValueByName(fieldName);
          const value = this._extractDisplayValue(rawValue);
          return this._escapeCsvField(value);
        });
        csvRows.push(values.join(','));
      }

      const csvContent = csvRows.join('\r\n');
      const blob = new Blob(['\uFEFF' + csvContent], { type: 'text/csv;charset=utf-8;' });
      const url = URL.createObjectURL(blob);
      const link = document.createElement('a');
      const timestamp = new Date().toISOString().slice(0, 10);

      link.setAttribute('href', url);
      link.setAttribute('download', `export-${timestamp}.csv`);
      link.style.visibility = 'hidden';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);

      URL.revokeObjectURL(url);

      Log.info(LOG_SOURCE, `Exported ${selectedRows.length} items to CSV`);
    } catch (error) {
      Log.error(LOG_SOURCE, error as Error);
      Dialog.alert(`Error exporting to CSV: ${(error as Error).message}`);
    }
  }

  /**
   * Opens the AssignPanel to bulk-assign selected items to a person.
   * Dynamically imports the panel component to reduce initial bundle size.
   */
  private async _handleAssignTo(selectedRows: readonly RowAccessor[]): Promise<void> {
    try {
      const itemIds = selectedRows.map(row =>
        parseInt(row.getValueByName('ID'), 10)
      );

      const component = await import(
        /* webpackChunkName: 'assign-panel' */
        './components/AssignPanel'
      );

      const element = document.createElement('div');
      document.body.appendChild(element);

      const React = await import('react');
      const ReactDOM = await import('react-dom');

      const onDismiss = (): void => {
        ReactDOM.unmountComponentAtNode(element);
        document.body.removeChild(element);
      };

      const onAssign = async (userId: number): Promise<void> => {
        const listId = this.context.listView.list.guid.toString();
        const [batchedSP, execute] = this._sp.batched();
        const list = batchedSP.web.lists.getById(listId);

        for (const itemId of itemIds) {
          list.items.getById(itemId).update({
            AssignedToId: userId
          });
        }

        await execute();

        await Dialog.alert(
          `Successfully assigned ${itemIds.length} item(s). Refresh the page to see updates.`
        );

        onDismiss();
      };

      ReactDOM.render(
        React.createElement(component.AssignPanel, {
          context: this.context,
          isOpen: true,
          itemCount: itemIds.length,
          onDismiss: onDismiss,
          onAssign: onAssign
        }),
        element
      );
    } catch (error) {
      Log.error(LOG_SOURCE, error as Error);
      await Dialog.alert(
        `Error opening assignment panel: ${(error as Error).message}`
      );
    }
  }

  /**
   * Displays a confirmation dialog and returns the user's choice.
   */
  private async _confirm(message: string): Promise<boolean> {
    return new Promise<boolean>((resolve) => {
      Dialog.alert(message).then(() => resolve(true), () => resolve(false));
    });
  }

  /**
   * Escapes a value for safe inclusion in a CSV field.
   */
  private _escapeCsvField(value: string): string {
    if (!value) return '""';
    const escaped = value.replace(/"/g, '""');
    return `"${escaped}"`;
  }

  /**
   * Extracts a display-friendly string from a list item field value,
   * handling person fields, lookups, and other complex types.
   */
  private _extractDisplayValue(rawValue: unknown): string {
    if (rawValue === null || rawValue === undefined) {
      return '';
    }

    if (typeof rawValue === 'string') {
      return rawValue;
    }

    if (typeof rawValue === 'number' || typeof rawValue === 'boolean') {
      return String(rawValue);
    }

    if (Array.isArray(rawValue)) {
      return rawValue.map(v => this._extractDisplayValue(v)).join('; ');
    }

    if (typeof rawValue === 'object') {
      const obj = rawValue as Record<string, unknown>;
      // Person or lookup field
      if (obj.title) return String(obj.title);
      if (obj.label) return String(obj.label);
      if (obj.lookupValue) return String(obj.lookupValue);
      if (obj.email) return String(obj.email);
      return JSON.stringify(rawValue);
    }

    return String(rawValue);
  }
}
