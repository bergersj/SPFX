import * as React from 'react';
import './QuickLinks.module.scss';
import { IQuickLinksProps } from './IQuickLinksProps';
//import { escape } from '@microsoft/sp-lodash-subset';
import { Icon } from 'office-ui-fabric-react';
import Heading from '../../../components/heading/Heading';
import { isMobile } from "react-device-detect";

export interface IQuickLink {
  iconName: string;
  openInNewTab?: boolean;
  text: string;
  url: string;
}

export default class QuickLinks extends React.Component<IQuickLinksProps, {}> {
  public render(): React.ReactElement<IQuickLinksProps> {
    const {
      heading,
      items
    } = this.props;
      
    return (
      <div className="width100">
        <div className="linkHeading"><Heading heading={heading}></Heading></div>
        {
          <div className="quick-links">
            {
              items && items.map((item:IQuickLink) =>
                isMobile && item.url.indexOf("keynet.keybank") !== -1 ?
                  //use mobile redirect url if device is mobile and url links to old keynet
                    <a
                    data-interception="off"
                    href="https://kbna.sharepoint.com/Sitepages/mobile-redirect.aspx"
                    target={item.openInNewTab ? "_blank" : "_self"}
                    >
                      <div className="icon">
                          <Icon className="iconImage" iconName={item.iconName} />
                      </div>
                      <div className="title">{item.text}</div>
                    </a>
                  :
                    <a
                    data-interception="off"
                    href={item.url}
                    target={item.openInNewTab ? "_blank" : "_self"}
                    >
                      <div className="icon">
                          <Icon className="iconImage" iconName={item.iconName} />
                      </div>
                      <div className="title">{item.text}</div>
                    </a>
              )
            }
            {
              !items || items.length == 0 && 
              <div>No data exists at this time.</div>
            }
          </div>
        }
      </div>
    );

  }
}
