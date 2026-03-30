import { override } from '@microsoft/decorators';
import { Log } from '@microsoft/sp-core-library';
import {
  BaseFieldCustomizer,
  IFieldCustomizerCellEventParameters
} from '@microsoft/sp-listview-extensibility';
import * as React from 'react';
import * as ReactDOM from 'react-dom';

import { StatusBadge, IStatusBadgeProps } from './components/StatusBadge';

const LOG_SOURCE: string = 'StatusFieldCustomizer';

export interface IStatusFieldCustomizerProperties {
  // Reserved for future configuration
}

export default class StatusFieldCustomizer extends BaseFieldCustomizer<IStatusFieldCustomizerProperties> {
  @override
  public onInit(): Promise<void> {
    Log.info(LOG_SOURCE, 'Initialized StatusFieldCustomizer');
    return Promise.resolve();
  }

  @override
  public onRenderCell(event: IFieldCustomizerCellEventParameters): void {
    const statusValue: string = event.fieldValue?.toString() ?? '';

    const element = React.createElement<IStatusBadgeProps>(StatusBadge, {
      status: statusValue
    });

    ReactDOM.render(element, event.domElement);
  }

  @override
  public onDisposeCell(event: IFieldCustomizerCellEventParameters): void {
    ReactDOM.unmountComponentAtNode(event.domElement);
    super.onDisposeCell(event);
  }
}
