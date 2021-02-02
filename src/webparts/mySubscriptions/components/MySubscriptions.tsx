import * as React from 'react';
import './MySubscriptions.module.scss';
import { IMySubscriptionsProps, ITermToggle } from './IMySubscriptionsProps';
import Heading from '../../../components/heading/Heading';
import { Toggle } from 'office-ui-fabric-react';
import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { sp } from "@pnp/sp";
import "@pnp/sp/profiles";

export interface MySubscriptionsState {
  userAccountName: string;
  userTopicSubscriptions: any;
  userCommunitySubscriptions: any;
}

export interface ITermData {
  name: string;
  isChecked: boolean;
}

export default class MySubscriptions extends React.Component<IMySubscriptionsProps,MySubscriptionsState, {}> {
  constructor(props: IMySubscriptionsProps) {
    super(props);
    this.state = {
      userAccountName: '',
      userTopicSubscriptions: '',
      userCommunitySubscriptions: '',
    };
  }
  
  public componentDidMount() {
    this.GetUserProperties();
  }

  private async GetUserProperties() {
    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      }
    });

    const response = await sp.profiles.myProperties.get();
    const topicSubscriptions = response.UserProfileProperties.find(cell => cell.Key === "SubscribedTopics").Value;
    const communitiesSubscriptions = response.UserProfileProperties.find(cell => cell.Key === "SubscribedCommunities").Value;
    const accountName = response.UserProfileProperties.find(cell => cell.Key === "AccountName").Value;

    this.setState({
      userAccountName: accountName,
      userTopicSubscriptions: topicSubscriptions,
      userCommunitySubscriptions: communitiesSubscriptions,
    });
  }

  public _onToggleChange(checked: boolean, term, isTopic: boolean) {
    const { 
      userTopicSubscriptions,
      userCommunitySubscriptions
     } = this.state;

     const userSubscriptions = isTopic ? userTopicSubscriptions : userCommunitySubscriptions;

    if(checked) {
      // add to subscription string
      const addNewSubscription = userSubscriptions.concat("|" + term.name);
      isTopic ? this.setState({ userTopicSubscriptions: addNewSubscription }, () => this.updateUserProfilePropertyValues(isTopic)) : this.setState({ userCommunitySubscriptions: addNewSubscription }, () => this.updateUserProfilePropertyValues(isTopic));
    } else {
      //remove from subscription string 
      let removeFilter = "";
      if(userSubscriptions.toLowerCase().includes("|" + term.name.toLowerCase())) {
        removeFilter = "|" + term.name.toLowerCase();
      } else {
        // first item in the subscription string
        removeFilter = term.name.toLowerCase();
      }

      const removeSubscription = userSubscriptions.toLowerCase().replace(removeFilter, '');
      isTopic? this.setState({userTopicSubscriptions: removeSubscription }, () => this.updateUserProfilePropertyValues(isTopic)) : this.setState({userCommunitySubscriptions: removeSubscription }, () => this.updateUserProfilePropertyValues(isTopic));
    }
  }

  private updateUserProfilePropertyValues = (isTopic) => {
    const { 
      userAccountName,
      userTopicSubscriptions,
      userCommunitySubscriptions
    } = this.state;

    const apiUrl = this.props.siteUrl + "/_api/SP.UserProfiles.PeopleManager/SetMultiValuedProfileProperty";
    const userData = {
      'accountName': "i:0#.f|membership|"+this.props.userEmail, //userAccountName,
      'propertyName': isTopic ? "SubscribedTopics" : "SubscribedCommunities",
      'propertyValues': isTopic ? userTopicSubscriptions.split("|") : userCommunitySubscriptions.split("|"),
    };
    const spOpts = {
      headers: {
        'Accept': 'application/json;odata=nometadata',
        'Content-type': 'application/json;odata=verbose',
        'odata-version': '3.0',
      },
      body: JSON.stringify(userData)
    };
    this.props.spHttpClient.post(apiUrl, SPHttpClient.configurations.v1, spOpts)
    .then((response: SPHttpClientResponse) => this.GetUserProperties());
  }

  public render(): React.ReactElement<IMySubscriptionsProps> {
    const {
      buttonText,
      buttonUrl,
      showButton,
      description,
      showDescription,
      title,
      topicTerms,
      communityTerms
    } = this.props;

    let topicToggleArray: ITermToggle[] = [];
    let topicNamesArray = typeof(this.props.topicTerms) == "undefined" ? [] : this.props.topicTerms.map(a=>a.name);
    const topicData = topicNamesArray.map((item) => {
      const {
        userTopicSubscriptions,
      } = this.state;
      
      const displayName = item.length > 17 ? item.substring(0,14).concat("...") : item; 

      topicToggleArray.push(
        {
          name: item,
          toolTip: displayName,
          isChecked: !!userTopicSubscriptions.split("|").find(topic => topic.toLowerCase() == item.toLowerCase()), 
        }
      );
      return topicToggleArray.sort((a, b) => (a.name > b.name) ? 1 : -1);
    });

    let communityToggleArray: ITermToggle[] = [];
    let communityNamesArray =  typeof(this.props.communityTerms) == "undefined" ? [] : this.props.communityTerms.map(a=>a.name);
    const communityData = communityNamesArray.map((item) => {
      const {
        userCommunitySubscriptions,
      } = this.state;

      const displayName = item.length > 17 ? item.substring(0,14).concat("...") : item; 

      communityToggleArray.push(
        {
          name: item,
          toolTip: displayName,
          isChecked: !!userCommunitySubscriptions.split("|").find(community => community.toLowerCase() == item.toLowerCase()), 
        }
      );
      return communityToggleArray.sort((a, b) => (a.name > b.name) ? 1 : -1);
    });

    return (
      <div className="mySubscriptions">
        <Heading heading={title} />
        {showButton &&
          <a 
            className="button"
            data-interception="off"
            href={buttonUrl}
          >
            {buttonText}
          </a>
        }
        {showDescription &&
        <div className="desc">
          {description}
        </div>
        }
        <Heading heading="Topics" />
        <div className="toggleContainer">
          {topicNamesArray.length > 0 ?
            topicToggleArray && topicToggleArray.map((topic: ITermToggle) =>(
              <div className="toggle">
                <div className="toggleLabel tooltip">{topic.toolTip}
                  {topic.name.length > 17 && <span className="tooltiptext">{topic.name}</span>}
                </div>
                <Toggle
                  className=""
                  checked= {topic.isChecked}
                  onChanged={(e) => this._onToggleChange(e, topic, true)}
                />
              </div>
            ))
            :
            <div>No topic subscriptions exist at this time.</div>
          }
        </div>
        <br/>  
        <Heading heading="Teams" />
        <div className="toggleContainer">
          {communityNamesArray.length > 0 ?
            communityToggleArray && communityToggleArray.map((community: ITermToggle) =>(
              <div className="toggle">
                <div className="toggleLabel tooltip">{community.toolTip}
                  {community.name.length > 17 && <span className="tooltiptext">{community.name}</span>}
                </div>
                <Toggle
                  className=""
                  checked= {community.isChecked}
                  onChanged={(e) => this._onToggleChange(e, community, false)}
                />
              </div>
            ))
            :
            <div>No team subscriptions exist at this time.</div>
          }
        </div>
      </div>
    );
  }
}
