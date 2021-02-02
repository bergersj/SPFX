import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneToggle
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'MySubscriptionsWebPartStrings';
import MySubscriptions from './components/MySubscriptions';
import { IMySubscriptionsProps } from './components/IMySubscriptionsProps';
import { PropertyFieldTermPicker } from '@pnp/spfx-property-controls/lib/PropertyFieldTermPicker';
import { IPickerTerms } from "@pnp/spfx-property-controls/lib/PropertyFieldTermPicker";

export interface IPropertyControlsTestWebPartProps {
  terms: IPickerTerms;
}

export interface IMySubscriptionsWebPartProps {
  title: string;
  buttonText: string;
  showButton: boolean;
  buttonUrl: string;
  description: string;
  showDescription: boolean;
  topicTerms: [];
  communityTerms: [];
}

export default class MySubscriptionsWebPart extends BaseClientSideWebPart <IMySubscriptionsWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IMySubscriptionsProps> = React.createElement(
      MySubscriptions,
      {
        description: this.properties.description,
        showDescription: this.properties.showDescription,
        buttonText: this.properties.buttonText,
        showButton: this.properties.showButton,
        buttonUrl: this.properties.buttonUrl,
        siteUrl: this.context.pageContext.web.absoluteUrl,
        userEmail: this.context.pageContext.user.email,
        spHttpClient: this.context.spHttpClient,
        title: this.properties.title,
        topicTerms: this.properties.topicTerms,
        communityTerms: this.properties.communityTerms
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
                  label: "Webpart title"
                }),
                PropertyPaneTextField('buttonText', {
                  label: "Button Text"
                }),
                PropertyPaneTextField('buttonUrl', {
                  label: "Button Url"
                }),
                PropertyPaneToggle('showButton', {
                  label: "Display Button",
                  onText: "Show Button",
                  offText:"Hide Button"
                }),
                PropertyPaneTextField('description', {
                  label: "Description"
                }),
                PropertyPaneToggle('showDescription', {
                  label: "Display Description",
                  onText: "Show Description",
                  offText:"Hide Description"
                }),
                PropertyFieldTermPicker('topicTerms', {
                  label: 'Select Topic terms',
                  panelTitle: 'Select Topic terms',
                  initialValues: this.properties.topicTerms,
                  allowMultipleSelections: true,
                  excludeSystemGroup: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  //limitByGroupNameOrID: 'SPARK',
                  limitByTermsetNameOrID: 'Topic',
                  key: 'termSetsPickerFieldId'
                }),
                PropertyFieldTermPicker('communityTerms', {
                  label: 'Select Team terms',
                  panelTitle: 'Select Team terms',
                  initialValues: this.properties.communityTerms,
                  allowMultipleSelections: true,
                  excludeSystemGroup: false,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  context: this.context,
                  onGetErrorMessage: null,
                  deferredValidationTime: 0,
                  //limitByGroupNameOrID: 'SPARK',
                  limitByTermsetNameOrID: 'Team',
                  key: 'termSetsPickerFieldId'
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
