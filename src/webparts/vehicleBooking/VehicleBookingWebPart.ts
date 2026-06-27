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
  apiBaseUrl: string;  // cấu hình per-tenant qua Property Pane
}

export default class VehicleBookingWebPart extends BaseClientSideWebPart<IVehicleBookingWebPartProps> {

  public render(): void {
    const element = React.createElement(
      VehicleBookingApp,
      { context: this.context, apiBaseUrl: this.properties.apiBaseUrl || '' }
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
                PropertyPaneTextField('apiBaseUrl', {
                  label: 'API Base URL',
                  placeholder: 'https://api.yourcompany.com/api',
                  description: 'Để trống nếu chỉ dùng SharePoint lists',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
