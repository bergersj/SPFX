import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MyNewsWebPartStrings';
import MyNews from './components/MyNews';
import { IMyNewsProps, IMyNewsItem } from './components/IMyNewsProps';

export interface IMyNewsWebPartProps {
  title: string;
  siteUrl: string;
  actionText: string;
  actionUrl: string;
  actionTextLeft: string;
  actionUrlLeft: string;
  siteID: string;
}

export default class MyNewsWebPart extends BaseClientSideWebPart <IMyNewsWebPartProps> {


  public async render(): Promise<void> {

    const element: React.ReactElement<IMyNewsProps> = React.createElement(
      MyNews,
      {
        title: this.properties.title,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        actionText: this.properties.actionText,
        actionUrl: this.properties.actionUrl,
        actionTextLeft: this.properties.actionTextLeft,
        actionUrlLeft: this.properties.actionUrlLeft,
        context: this.context,
        siteID: this.context.pageContext.web.id,
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
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('actionText', {
                  label: "Right Link Text"
                }),
                PropertyPaneTextField('actionUrl', {
                  label: "Right Link URL"
                }),
                PropertyPaneTextField('actionTextLeft', {
                  label: "Left Link Text"
                }),
                PropertyPaneTextField('actionUrlLeft', {
                  label: "Left Link URL"
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
