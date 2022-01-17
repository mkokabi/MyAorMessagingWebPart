import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { AadHttpClient, HttpClientResponse } from '@microsoft/sp-http';

import * as strings from 'MyAorMessageWebPartStrings';
import { IMyAorMessageProps } from './components/IMyAorMessageProps';
import { MyAorMessage } from './components/MyAorMessage';

export interface IMyAorMessageWebPartProps {
  description: string;
}

export default class MyAorMessageWebPart extends BaseClientSideWebPart<IMyAorMessageWebPartProps> {
  private messagingClient: AadHttpClient;

  protected onInit(): Promise<void> {
    return new Promise<void>((resolve: () => void, reject: (error: any) => void): void => {
      this.context.aadHttpClientFactory
        .getClient('a3960e3e-11b9-4310-90a7-a9212ecbbf9b')
        .then((client: AadHttpClient): void => {
          this.messagingClient = client;
          resolve();
        }, err => reject(err));
    });
  }

  public render(): void {
    const element: React.ReactElement<IMyAorMessageProps> = React.createElement(
      MyAorMessage,
      {
        description: this.properties.description,
        messagingClient: this.messagingClient
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
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
