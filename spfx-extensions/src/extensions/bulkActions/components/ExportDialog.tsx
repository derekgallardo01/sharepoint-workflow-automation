import * as React from 'react';
import {
  Dialog,
  DialogType,
  DialogFooter,
  PrimaryButton,
  DefaultButton,
  Dropdown,
  IDropdownOption,
  Checkbox,
  DatePicker,
  Stack,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  Text,
  Label,
  Separator,
  IStackTokens
} from '@fluentui/react';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

export type ExportFormat = 'csv' | 'excel' | 'json';

export interface IExportColumn {
  key: string;
  name: string;
  fieldType: string;
  selected: boolean;
}

export interface IExportDialogProps {
  context: ListViewCommandSetContext;
  isOpen: boolean;
  listName: string;
  availableColumns: IExportColumn[];
  totalItemCount: number;
  onDismiss: () => void;
  onExport: (options: IExportOptions) => Promise<void>;
}

export interface IExportOptions {
  format: ExportFormat;
  selectedColumns: string[];
  dateFrom: Date | null;
  dateTo: Date | null;
  includeMetadata: boolean;
}

interface IExportDialogState {
  selectedFormat: ExportFormat;
  columns: IExportColumn[];
  dateFrom: Date | null;
  dateTo: Date | null;
  includeMetadata: boolean;
  isExporting: boolean;
  exportProgress: number;
  error: string | null;
  success: boolean;
}

const stackTokens: IStackTokens = { childrenGap: 12 };

const formatOptions: IDropdownOption[] = [
  { key: 'csv', text: 'CSV (.csv)', data: { icon: 'TextDocument' } },
  { key: 'excel', text: 'Excel (.xlsx)', data: { icon: 'ExcelDocument' } },
  { key: 'json', text: 'JSON (.json)', data: { icon: 'Code' } }
];

export class ExportDialog extends React.Component<IExportDialogProps, IExportDialogState> {
  private _exportIntervalId: ReturnType<typeof setInterval> | null = null;

  constructor(props: IExportDialogProps) {
    super(props);
    this.state = {
      selectedFormat: 'csv',
      columns: props.availableColumns.map((col) => ({ ...col, selected: true })),
      dateFrom: null,
      dateTo: null,
      includeMetadata: false,
      isExporting: false,
      exportProgress: 0,
      error: null,
      success: false
    };
  }

  public componentDidUpdate(prevProps: IExportDialogProps): void {
    if (prevProps.availableColumns !== this.props.availableColumns) {
      this.setState({
        columns: this.props.availableColumns.map((col) => ({ ...col, selected: true }))
      });
    }
    // Reset state when dialog opens
    if (!prevProps.isOpen && this.props.isOpen) {
      this.setState({
        isExporting: false,
        exportProgress: 0,
        error: null,
        success: false
      });
    }
  }

  public componentWillUnmount(): void {
    if (this._exportIntervalId) {
      clearInterval(this._exportIntervalId);
    }
  }

  public render(): React.ReactElement<IExportDialogProps> {
    const { isOpen, onDismiss, listName, totalItemCount } = this.props;
    const {
      selectedFormat,
      columns,
      dateFrom,
      dateTo,
      includeMetadata,
      isExporting,
      exportProgress,
      error,
      success
    } = this.state;

    const selectedColumnCount = columns.filter((c) => c.selected).length;
    const allSelected = selectedColumnCount === columns.length;
    const noneSelected = selectedColumnCount === 0;

    return (
      <Dialog
        hidden={!isOpen}
        onDismiss={isExporting ? undefined : onDismiss}
        dialogContentProps={{
          type: DialogType.largeHeader,
          title: 'Export List Data',
          subText: `Export items from "${listName}" (${totalItemCount} item${totalItemCount !== 1 ? 's' : ''})`
        }}
        modalProps={{
          isBlocking: isExporting,
          styles: { main: { minWidth: 520, maxWidth: 640 } }
        }}
      >
        <Stack tokens={stackTokens}>
          {error && (
            <MessageBar
              messageBarType={MessageBarType.error}
              onDismiss={() => this.setState({ error: null })}
              dismissButtonAriaLabel="Close"
            >
              {error}
            </MessageBar>
          )}

          {success && (
            <MessageBar messageBarType={MessageBarType.success}>
              Export completed successfully. Your file download should begin automatically.
            </MessageBar>
          )}

          {isExporting && (
            <ProgressIndicator
              label={`Exporting as ${selectedFormat.toUpperCase()}...`}
              description={`${Math.round(exportProgress * 100)}% complete`}
              percentComplete={exportProgress}
            />
          )}

          {/* Export Format */}
          <Dropdown
            label="Export Format"
            selectedKey={selectedFormat}
            options={formatOptions}
            onChange={this._onFormatChange}
            disabled={isExporting}
            required
          />

          <Separator />

          {/* Column Selection */}
          <Stack tokens={{ childrenGap: 8 }}>
            <Stack horizontal horizontalAlign="space-between" verticalAlign="center">
              <Label>Columns to Export</Label>
              <Text variant="small" styles={{ root: { color: '#666' } }}>
                {selectedColumnCount} of {columns.length} selected
              </Text>
            </Stack>

            <Stack horizontal tokens={{ childrenGap: 12 }}>
              <DefaultButton
                text="Select All"
                onClick={this._selectAllColumns}
                disabled={isExporting || allSelected}
                styles={{ root: { minWidth: 0, padding: '0 8px' } }}
              />
              <DefaultButton
                text="Deselect All"
                onClick={this._deselectAllColumns}
                disabled={isExporting || noneSelected}
                styles={{ root: { minWidth: 0, padding: '0 8px' } }}
              />
            </Stack>

            <Stack
              tokens={{ childrenGap: 6 }}
              styles={{
                root: {
                  maxHeight: 200,
                  overflowY: 'auto',
                  border: '1px solid #edebe9',
                  borderRadius: 2,
                  padding: 8
                }
              }}
            >
              {columns.map((col, index) => (
                <Checkbox
                  key={col.key}
                  label={`${col.name} (${col.fieldType})`}
                  checked={col.selected}
                  onChange={(_ev, checked) => this._onColumnToggle(index, checked || false)}
                  disabled={isExporting}
                />
              ))}
            </Stack>
          </Stack>

          <Separator />

          {/* Date Range Filter */}
          <Stack tokens={{ childrenGap: 8 }}>
            <Label>Date Range Filter (Optional)</Label>
            <Text variant="small" styles={{ root: { color: '#666' } }}>
              Filter items by their Modified date before exporting.
            </Text>
            <Stack horizontal tokens={{ childrenGap: 12 }}>
              <DatePicker
                label="From"
                value={dateFrom || undefined}
                onSelectDate={this._onDateFromChange}
                placeholder="Start date..."
                disabled={isExporting}
                styles={{ root: { flex: 1 } }}
                allowTextInput
              />
              <DatePicker
                label="To"
                value={dateTo || undefined}
                onSelectDate={this._onDateToChange}
                placeholder="End date..."
                disabled={isExporting}
                styles={{ root: { flex: 1 } }}
                minDate={dateFrom || undefined}
                allowTextInput
              />
            </Stack>
          </Stack>

          {/* Include Metadata */}
          <Checkbox
            label="Include item metadata (ID, Created, Modified, Author)"
            checked={includeMetadata}
            onChange={(_ev, checked) => this.setState({ includeMetadata: checked || false })}
            disabled={isExporting}
          />
        </Stack>

        <DialogFooter>
          <PrimaryButton
            text={isExporting ? 'Exporting...' : 'Export'}
            onClick={this._onExportClick}
            disabled={isExporting || noneSelected}
            iconProps={{ iconName: 'Download' }}
          />
          <DefaultButton
            text="Cancel"
            onClick={onDismiss}
            disabled={isExporting}
          />
        </DialogFooter>
      </Dialog>
    );
  }

  // ---------------------------------------------------------------------------
  // Event Handlers
  // ---------------------------------------------------------------------------

  private _onFormatChange = (
    _event: React.FormEvent<HTMLDivElement>,
    option?: IDropdownOption
  ): void => {
    if (option) {
      this.setState({ selectedFormat: option.key as ExportFormat });
    }
  };

  private _onColumnToggle = (index: number, checked: boolean): void => {
    this.setState((prevState) => {
      const columns = [...prevState.columns];
      columns[index] = { ...columns[index], selected: checked };
      return { columns };
    });
  };

  private _selectAllColumns = (): void => {
    this.setState((prevState) => ({
      columns: prevState.columns.map((col) => ({ ...col, selected: true }))
    }));
  };

  private _deselectAllColumns = (): void => {
    this.setState((prevState) => ({
      columns: prevState.columns.map((col) => ({ ...col, selected: false }))
    }));
  };

  private _onDateFromChange = (date: Date | null | undefined): void => {
    this.setState({ dateFrom: date || null });
  };

  private _onDateToChange = (date: Date | null | undefined): void => {
    this.setState({ dateTo: date || null });
  };

  private _onExportClick = async (): Promise<void> => {
    const { onExport } = this.props;
    const { selectedFormat, columns, dateFrom, dateTo, includeMetadata } = this.state;

    const selectedColumns = columns
      .filter((col) => col.selected)
      .map((col) => col.key);

    if (selectedColumns.length === 0) {
      this.setState({ error: 'Please select at least one column to export.' });
      return;
    }

    this.setState({
      isExporting: true,
      exportProgress: 0,
      error: null,
      success: false
    });

    // Simulate progress updates (the actual export happens in the parent)
    this._exportIntervalId = setInterval(() => {
      this.setState((prevState) => {
        const newProgress = Math.min(prevState.exportProgress + 0.05, 0.95);
        return { exportProgress: newProgress };
      });
    }, 200);

    try {
      await onExport({
        format: selectedFormat,
        selectedColumns,
        dateFrom,
        dateTo,
        includeMetadata
      });

      if (this._exportIntervalId) {
        clearInterval(this._exportIntervalId);
        this._exportIntervalId = null;
      }

      this.setState({
        isExporting: false,
        exportProgress: 1,
        success: true
      });
    } catch (err) {
      if (this._exportIntervalId) {
        clearInterval(this._exportIntervalId);
        this._exportIntervalId = null;
      }

      const message =
        err instanceof Error ? err.message : 'An unexpected error occurred during export.';

      this.setState({
        isExporting: false,
        exportProgress: 0,
        error: message
      });
    }
  };
}
