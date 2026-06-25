import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import RequestManagement from './components/RequestManagement';
import { IRequestManagementProps } from './components/IRequestManagementProps';

export interface IRequestManagementWebPartProps {
  listName: string;
}

export default class RequestManagementWebPart extends BaseClientSideWebPart<IRequestManagementWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IRequestManagementProps> = React.createElement(
      RequestManagement,
      {
        context: this.context,
        listName: this.properties.listName || 'Requests',
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
          header: { description: 'Configure Request Management Web Part' },
          groups: [
            {
              groupName: 'Settings',
              groupFields: [
                PropertyPaneTextField('listName', {
                  label: 'SharePoint List Name',
                  placeholder: 'e.g. Requests',
                  description: 'Enter the exact name of the SharePoint list containing requests'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
