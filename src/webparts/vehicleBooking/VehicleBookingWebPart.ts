import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import VehicleBookingApp, { IVehicleBookingAppProps } from './components/VehicleBookingApp';

export interface IVehicleBookingWebPartProps {
  description: string;
}

export default class VehicleBookingWebPart extends BaseClientSideWebPart<IVehicleBookingWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IVehicleBookingAppProps> = React.createElement(
      VehicleBookingApp,
      { context: this.context }
    );
    ReactDom.render(element, this.domElement);
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: { description: 'Vehicle Booking System' },
          groups: [
            {
              groupName: 'Settings',
              groupFields: [
                PropertyPaneTextField('description', {
                  label: 'Description',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
