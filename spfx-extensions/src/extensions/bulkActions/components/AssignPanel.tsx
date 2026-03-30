import * as React from 'react';
import {
  Panel,
  PanelType,
  PrimaryButton,
  DefaultButton,
  Stack,
  MessageBar,
  MessageBarType,
  Spinner,
  SpinnerSize,
  Text
} from '@fluentui/react';
import { PeoplePicker, PrincipalType } from '@pnp/spfx-controls-react/lib/PeoplePicker';
import { ListViewCommandSetContext } from '@microsoft/sp-listview-extensibility';

export interface IAssignPanelProps {
  context: ListViewCommandSetContext;
  isOpen: boolean;
  itemCount: number;
  onDismiss: () => void;
  onAssign: (userId: number) => Promise<void>;
}

interface IAssignPanelState {
  selectedUserId: number | null;
  selectedUserName: string;
  isProcessing: boolean;
  error: string | null;
}

export class AssignPanel extends React.Component<IAssignPanelProps, IAssignPanelState> {
  constructor(props: IAssignPanelProps) {
    super(props);
    this.state = {
      selectedUserId: null,
      selectedUserName: '',
      isProcessing: false,
      error: null
    };
  }

  public render(): React.ReactElement<IAssignPanelProps> {
    const { isOpen, itemCount, onDismiss } = this.props;
    const { selectedUserId, isProcessing, error } = this.state;

    return (
      <Panel
        isOpen={isOpen}
        type={PanelType.medium}
        headerText="Assign Items"
        onDismiss={onDismiss}
        isFooterAtBottom={true}
        onRenderFooterContent={this._onRenderFooterContent}
        closeButtonAriaLabel="Close"
        isBlocking={isProcessing}
      >
        <Stack tokens={{ childrenGap: 16, padding: '16px 0' }}>
          <Text variant="medium">
            Select a person to assign to {itemCount} selected item(s).
          </Text>

          {error && (
            <MessageBar
              messageBarType={MessageBarType.error}
              onDismiss={() => this.setState({ error: null })}
              dismissButtonAriaLabel="Close"
            >
              {error}
            </MessageBar>
          )}

          <PeoplePicker
            context={this.props.context as never}
            titleText="Assign to"
            personSelectionLimit={1}
            required={true}
            showHiddenInUI={false}
            principalTypes={[PrincipalType.User]}
            resolveDelay={300}
            onChange={this._onPeoplePickerChange}
            placeholder="Search for a person..."
            disabled={isProcessing}
          />

          {isProcessing && (
            <Stack horizontal tokens={{ childrenGap: 8 }} verticalAlign="center">
              <Spinner size={SpinnerSize.small} />
              <Text variant="small">
                Assigning {itemCount} item(s)...
              </Text>
            </Stack>
          )}

          {selectedUserId && !isProcessing && (
            <MessageBar messageBarType={MessageBarType.info}>
              {this.state.selectedUserName} will be assigned to {itemCount} item(s).
            </MessageBar>
          )}
        </Stack>
      </Panel>
    );
  }

  private _onRenderFooterContent = (): React.ReactElement => {
    const { onDismiss } = this.props;
    const { selectedUserId, isProcessing } = this.state;

    return (
      <Stack horizontal tokens={{ childrenGap: 8 }}>
        <PrimaryButton
          text="Assign"
          onClick={this._onAssignClick}
          disabled={!selectedUserId || isProcessing}
        />
        <DefaultButton
          text="Cancel"
          onClick={onDismiss}
          disabled={isProcessing}
        />
      </Stack>
    );
  }

  private _onPeoplePickerChange = (items: { id: string; text: string }[]): void => {
    if (items && items.length > 0) {
      this.setState({
        selectedUserId: parseInt(items[0].id, 10),
        selectedUserName: items[0].text,
        error: null
      });
    } else {
      this.setState({
        selectedUserId: null,
        selectedUserName: ''
      });
    }
  }

  private _onAssignClick = async (): Promise<void> => {
    const { selectedUserId } = this.state;
    if (!selectedUserId) return;

    this.setState({ isProcessing: true, error: null });

    try {
      await this.props.onAssign(selectedUserId);
    } catch (error) {
      this.setState({
        isProcessing: false,
        error: `Assignment failed: ${(error as Error).message}. Please try again.`
      });
    }
  }
}
