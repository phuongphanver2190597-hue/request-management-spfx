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
  apiBaseUrl: string;
  claudeApiKey: string;
}

export default class VehicleBookingWebPart extends BaseClientSideWebPart<IVehicleBookingWebPartProps> {

  public render(): void {
    const element = React.createElement(
      VehicleBookingApp,
      {
        context: this.context,
        apiBaseUrl: this.properties.apiBaseUrl || '',
        claudeApiKey: this.properties.claudeApiKey || '',
      }
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
            {
              groupName: 'Chatbot (Claude AI)',
              groupFields: [
                PropertyPaneTextField('claudeApiKey', {
                  label: 'Claude API Key',
                  placeholder: 'sk-ant-...',
                  description: 'Lấy tại console.anthropic.com — để trống để tắt chatbot',
                }),
              ],
            },
          ],
        },
      ],
    };
  }
}
