import * as React from 'react';
import {
  Panel,
  PanelType,
  PrimaryButton,
  DefaultButton,
  Stack,
  ProgressIndicator,
  MessageBar,
  MessageBarType,
  Icon,
  Text,
  DetailsList,
  DetailsListLayoutMode,
  IColumn,
  SelectionMode,
} from '@fluentui/react';

// ---------------------------------------------------------------------------
// Types
// ---------------------------------------------------------------------------

/** Status of an individual item in the batch operation. */
export type ItemOperationStatus = 'pending' | 'processing' | 'success' | 'failed';

/** Represents a single item being processed in a batch operation. */
export interface IBatchItem {
  /** SharePoint list item ID. */
  id: number;
  /** Display title of the item. */
  title: string;
  /** Current processing status. */
  status: ItemOperationStatus;
  /** Error message if status is 'failed'. */
  errorMessage?: string;
}

/** Summary of a completed batch operation. */
export interface IBatchSummary {
  total: number;
  succeeded: number;
  failed: number;
  durationMs: number;
}

/** Props for the ProgressPanel component. */
export interface IProgressPanelProps {
  /** Whether the panel is open. */
  isOpen: boolean;
  /** Title displayed in the panel header. */
  operationTitle: string;
  /** The items being processed. */
  items: IBatchItem[];
  /** Whether the operation is currently running. */
  isRunning: boolean;
  /** Completion summary (populated after operation finishes). */
  summary: IBatchSummary | null;
  /** Callback to dismiss the panel. */
  onDismiss: () => void;
  /** Callback to cancel the in-progress operation. */
  onCancel: () => void;
  /** Callback to retry a specific failed item. */
  onRetryItem: (itemId: number) => Promise<void>;
  /** Callback to retry all failed items. */
  onRetryAll: () => Promise<void>;
}

/** Internal state. */
interface IProgressPanelState {
  /** Set of item IDs currently being retried. */
  retryingItems: Set<number>;
  /** Whether a retry-all operation is in progress. */
  isRetryingAll: boolean;
}

// ---------------------------------------------------------------------------
// Helpers
// ---------------------------------------------------------------------------

function getStatusIcon(status: ItemOperationStatus): { name: string; color: string } {
  switch (status) {
    case 'pending':
      return { name: 'Clock', color: '#605e5c' };
    case 'processing':
      return { name: 'Sync', color: '#0078d4' };
    case 'success':
      return { name: 'Completed', color: '#107c10' };
    case 'failed':
      return { name: 'ErrorBadge', color: '#d13438' };
    default:
      return { name: 'Unknown', color: '#605e5c' };
  }
}

function formatDuration(ms: number): string {
  if (ms < 1000) return `${ms}ms`;
  const seconds = Math.floor(ms / 1000);
  if (seconds < 60) return `${seconds}s`;
  const minutes = Math.floor(seconds / 60);
  const remainingSeconds = seconds % 60;
  return `${minutes}m ${remainingSeconds}s`;
}

// ---------------------------------------------------------------------------
// Component
// ---------------------------------------------------------------------------

/**
 * ProgressPanel displays real-time progress for batch operations on
 * SharePoint list items.
 *
 * Features:
 * - Progress bar with percentage and item count
 * - Item-by-item status (pending / processing / success / failed)
 * - Error details for failed items with per-item retry button
 * - Summary on completion
 * - Cancel button during operation
 */
export class ProgressPanel extends React.Component<IProgressPanelProps, IProgressPanelState> {
  constructor(props: IProgressPanelProps) {
    super(props);
    this.state = {
      retryingItems: new Set<number>(),
      isRetryingAll: false,
    };
  }

  // -----------------------------------------------------------------------
  // Computed values
  // -----------------------------------------------------------------------

  private get completedCount(): number {
    return this.props.items.filter(
      (i) => i.status === 'success' || i.status === 'failed'
    ).length;
  }

  private get progressPercent(): number {
    const total = this.props.items.length;
    return total > 0 ? this.completedCount / total : 0;
  }

  private get failedItems(): IBatchItem[] {
    return this.props.items.filter((i) => i.status === 'failed');
  }

  // -----------------------------------------------------------------------
  // Actions
  // -----------------------------------------------------------------------

  private handleRetryItem = async (itemId: number): Promise<void> => {
    this.setState((prev) => {
      const next = new Set(prev.retryingItems);
      next.add(itemId);
      return { retryingItems: next };
    });

    try {
      await this.props.onRetryItem(itemId);
    } finally {
      this.setState((prev) => {
        const next = new Set(prev.retryingItems);
        next.delete(itemId);
        return { retryingItems: next };
      });
    }
  };

  private handleRetryAll = async (): Promise<void> => {
    this.setState({ isRetryingAll: true });
    try {
      await this.props.onRetryAll();
    } finally {
      this.setState({ isRetryingAll: false });
    }
  };

  // -----------------------------------------------------------------------
  // Render helpers
  // -----------------------------------------------------------------------

  private renderProgressSection(): React.ReactElement {
    const { items, isRunning } = this.props;
    const total = items.length;
    const completed = this.completedCount;
    const percent = this.progressPercent;

    const description = isRunning
      ? `Processing ${completed} of ${total} items...`
      : `${completed} of ${total} items processed`;

    return (
      <Stack tokens={{ childrenGap: 8 }}>
        <ProgressIndicator
          label={this.props.operationTitle}
          description={description}
          percentComplete={percent}
          barHeight={4}
        />
        <Text variant="small" style={{ color: '#605e5c' }}>
          {Math.round(percent * 100)}% complete
        </Text>
      </Stack>
    );
  }

  private renderSummary(): React.ReactElement | null {
    const { summary } = this.props;
    if (!summary) return null;

    const allSucceeded = summary.failed === 0;
    const messageType = allSucceeded
      ? MessageBarType.success
      : MessageBarType.warning;

    return (
      <MessageBar messageBarType={messageType}>
        <strong>Operation Complete</strong>
        <br />
        {summary.succeeded} succeeded, {summary.failed} failed
        {' '}(completed in {formatDuration(summary.durationMs)})
      </MessageBar>
    );
  }

  private renderItemList(): React.ReactElement {
    const { items } = this.props;
    const { retryingItems } = this.state;

    const columns: IColumn[] = [
      {
        key: 'status',
        name: 'Status',
        minWidth: 40,
        maxWidth: 40,
        onRender: (item: IBatchItem) => {
          const icon = getStatusIcon(item.status);
          return (
            <Icon
              iconName={icon.name}
              style={{ color: icon.color, fontSize: 16 }}
              title={item.status}
              aria-label={item.status}
            />
          );
        },
      },
      {
        key: 'title',
        name: 'Item',
        minWidth: 120,
        maxWidth: 300,
        isResizable: true,
        onRender: (item: IBatchItem) => (
          <Text variant="small" title={item.title}>
            {item.title}
          </Text>
        ),
      },
      {
        key: 'detail',
        name: 'Detail',
        minWidth: 100,
        isResizable: true,
        onRender: (item: IBatchItem) => {
          if (item.status === 'failed' && item.errorMessage) {
            return (
              <Text
                variant="small"
                style={{ color: '#d13438' }}
                title={item.errorMessage}
              >
                {item.errorMessage}
              </Text>
            );
          }
          if (item.status === 'processing') {
            return (
              <Text variant="small" style={{ color: '#0078d4', fontStyle: 'italic' }}>
                Processing...
              </Text>
            );
          }
          return null;
        },
      },
      {
        key: 'actions',
        name: '',
        minWidth: 60,
        maxWidth: 60,
        onRender: (item: IBatchItem) => {
          if (item.status === 'failed') {
            const isRetrying = retryingItems.has(item.id);
            return (
              <DefaultButton
                text={isRetrying ? '...' : 'Retry'}
                onClick={() => this.handleRetryItem(item.id)}
                disabled={isRetrying}
                styles={{
                  root: { height: 24, minWidth: 50, padding: '0 8px', fontSize: 12 },
                }}
                aria-label={`Retry ${item.title}`}
              />
            );
          }
          return null;
        },
      },
    ];

    return (
      <DetailsList
        items={items}
        columns={columns}
        selectionMode={SelectionMode.none}
        layoutMode={DetailsListLayoutMode.justified}
        compact={true}
        isHeaderVisible={true}
        getKey={(item: IBatchItem) => String(item.id)}
        ariaLabelForGrid="Batch operation items"
      />
    );
  }

  private renderFooter = (): React.ReactElement => {
    const { isRunning, onCancel, onDismiss } = this.props;
    const { isRetryingAll } = this.state;
    const hasFailures = this.failedItems.length > 0;

    return (
      <Stack horizontal tokens={{ childrenGap: 8 }} style={{ padding: '16px 0' }}>
        {isRunning && (
          <DefaultButton
            text="Cancel"
            onClick={onCancel}
            iconProps={{ iconName: 'Cancel' }}
          />
        )}
        {!isRunning && hasFailures && (
          <PrimaryButton
            text={isRetryingAll ? 'Retrying...' : `Retry Failed (${this.failedItems.length})`}
            onClick={this.handleRetryAll}
            disabled={isRetryingAll}
            iconProps={{ iconName: 'Refresh' }}
          />
        )}
        {!isRunning && (
          <DefaultButton
            text="Close"
            onClick={onDismiss}
          />
        )}
      </Stack>
    );
  };

  // -----------------------------------------------------------------------
  // Render
  // -----------------------------------------------------------------------

  public render(): React.ReactElement<IProgressPanelProps> {
    const { isOpen, isRunning, onDismiss, summary } = this.props;

    return (
      <Panel
        isOpen={isOpen}
        type={PanelType.medium}
        headerText={this.props.operationTitle}
        onDismiss={onDismiss}
        isFooterAtBottom={true}
        onRenderFooterContent={this.renderFooter}
        closeButtonAriaLabel="Close progress panel"
        isBlocking={isRunning}
        isLightDismiss={!isRunning}
      >
        <Stack tokens={{ childrenGap: 16, padding: '16px 0' }}>
          {/* Progress bar */}
          {this.renderProgressSection()}

          {/* Completion summary */}
          {!isRunning && summary && this.renderSummary()}

          {/* Item-by-item list */}
          {this.renderItemList()}
        </Stack>
      </Panel>
    );
  }
}

export default ProgressPanel;
