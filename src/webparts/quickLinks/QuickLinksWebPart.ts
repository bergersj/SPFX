import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'QuickLinksWebPartStrings';
import QuickLinks from './components/QuickLinks';
import { IQuickLinksProps, ILink } from './components/IQuickLinksProps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';

export interface IQuickLinksWebPartProps {
  collectionData: ILink[];
  title: string;
}

export default class QuickLinksWebPart extends BaseClientSideWebPart <IQuickLinksWebPartProps> {

  public render(): void {
    const element: React.ReactElement<IQuickLinksProps> = React.createElement(
      QuickLinks,
      {
        heading: this.properties.title,
        items: this.properties.collectionData || [],
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
              groupFields: [
                PropertyPaneTextField('title', {
                  label: "Webpart title"
                }),
                PropertyFieldCollectionData("collectionData", {
                  key: "collectionData",
                  label: "Quick link data",
                  panelHeader: "Quick Links Data",
                  manageBtnLabel: "Manage Links",
                  value: this.properties.collectionData,
                  enableSorting: true,
                  fields: [
                    {
                      id: "text",
                      title: "Title",
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "url",
                      title: "Link Url",
                      type: CustomCollectionFieldType.url,
                      required: true
                    },
                    {
                      id: "iconName",
                      title: "Icon Name",
                      type: CustomCollectionFieldType.dropdown,
                      options: [
                        {
                          key:"Globe",
                          text:"Globe"
                        },
                        {
                          key:"Airplane",
                          text:"Airplane"
                        },
                        {
                          key:"AlertSolid",
                          text:"Alert"
                        },
                        {
                          key:"BarChart4",
                          text:"Bar Chart"
                        },
                        {
                          key:"Ringer",
                          text:"Bell"
                        },
                        {
                          key:"Calendar",
                          text:"Calendar"
                        },
                        {
                          key:"OfficeChat",
                          text:"Chat"
                        },
                        {
                          key:"ClipboardList",
                          text:"Clipboard List"
                        },
                        {
                          key:"Clock",
                          text:"Clock"
                        },
                        {
                          key:"CloudWeather",
                          text:"Cloud"
                        },
                        {
                          key:"CoffeeScript",
                          text:"Coffee"
                        },
                        {
                          key:"CompassNW",
                          text:"Compass"
                        },
                        {
                          key:"ContactCard",
                          text:"Contact Card"
                        },
                        {
                          key:"D365CustomerInsights",
                          text:"Customer Insights"
                        },  
                        {
                          key:"Mail",
                          text:"Envelope"
                        },
                        {
                          key:"FinancialSolid",
                          text:"Graph Trending Up"
                        },
                        {
                          key:"Commitments",
                          text:"Handshake"
                        },
                        {
                          key:"Health",
                          text:"Health"
                        },
                        {
                          key:"Headset",
                          text:"Help Desk"
                        },
                        {
                          key:"Info",
                          text:"Info"
                        },
                        {
                          key:"Link",
                          text:"Link"
                        },
                        {
                          key:"Settings",
                          text:"Settings"
                        },
                        {
                          key:"ShieldSolid",
                          text:"Shield"
                        },
                        {
                          key:"Emoji2",
                          text:"Smile"
                        },
                        {
                          key:"Savings",
                          text:"Piggy Bank"
                        },
                        {
                          key:"Pinned",
                          text:"Pin"
                        },
                        {
                          key:"Tag",
                          text:"Tag"
                        },
                        {
                          key:"DeveloperTools",
                          text:"Tools"
                        },
                        {
                          key:"Warning",
                          text:"Warning"
                        }
                      ],
                      required: true
                    },
                    {
                      id: "openInNewTab",
                      title: "Open in New Window",
                      type: CustomCollectionFieldType.boolean
                    }
                  ],
                  disabled: false
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
