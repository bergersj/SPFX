import { SPHttpClient } from '@microsoft/sp-http';


export interface IMySubscriptionsProps {
  description: string;
  showDescription: boolean;
  buttonText: string;
  buttonUrl: string;
  showButton: boolean;
  siteUrl: string;
  userEmail: string;
  spHttpClient: SPHttpClient;
  title: string;
  topicTerms:ITermToggle[];
  communityTerms:ITermToggle[];
}

export interface ITermToggle {
  name: string;
  toolTip: string;
  isChecked: boolean;
}
