import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClient } from '@microsoft/sp-http';
import * as MicrosoftGraph from '@microsoft/microsoft-graph-types';
import * as strings from 'LazyEmailViewerWebPartStrings';
import LazyEmailViewer from './components/LazyEmailViewer';
import { ILazyEmailViewerProps } from './components/ILazyEmailViewerProps';

export interface ILazyEmailViewerWebPartProps {
  title: string;
}

export default class LazyEmailViewerWebPart extends BaseClientSideWebPart<ILazyEmailViewerWebPartProps> {

  public render(): void {
    
    const element: React.ReactElement<ILazyEmailViewerProps> = React.createElement(
      LazyEmailViewer,
      {
        title: this.properties.title,
        mSGraphClientPromise: this.context.msGraphClientFactory.getClient(),
        httpClient:this.context.httpClient
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
                PropertyPaneTextField('title', {
                  label: 'App Title'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
