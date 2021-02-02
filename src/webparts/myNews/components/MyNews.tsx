import * as React from 'react';
import './MyNews.module.scss';
import { IMyNewsProps, IMyNewsItem } from './IMyNewsProps';
import { sp } from '@pnp/sp/presets/all';
import "@pnp/sp/profiles";
import { SPHttpClient } from '@microsoft/sp-http';
import  Heading, { ILink } from '../../../components/heading/Heading';
import * as moment from 'moment';
import { Spinner, SpinnerSize } from "office-ui-fabric-react/lib/Spinner";

export interface MyNewsState {
  newsArticles: IMyNewsItem[];
  isLoading: boolean;
}

export interface IUserSubscriptions {
  topicSubscriptions: [];
  communitySubscriptions: [];
}

export const MyNewsItemOne = (props: IMyNewsItem) => {
  let formatDate = (date: string) => {
    let dt = moment(date);
    return dt.format("D MMM YYYY");
  };

  return <div className="articleCardOne boxHover">
    <a
      href={props.url}
      target="_blank"
      data-interception="off"
    >
      { props.previewMediaUrl === null ?
        <div>No Image Available</div> 
        :
        <div>
          <img 
            src={props.previewMediaUrl}
            className="articleImage"
          />
        </div>
      }
    </a>
    <div className="articleInfo">
      <div className="articleDate">{formatDate(props.articleDate).toUpperCase()}</div>
      <div className="articleTitle">
        <a href={props.url} target="_blank" data-interceptio="off">
          {props.title}
        </a>
      </div>
      <div className="articleDescription">{props.articleDescription}</div>
    </div>
  </div>;
};

export const MyNewsItemTwo = (props: IMyNewsItem) => {
  let formatDate = (date: string) => {
    let dt = moment(date);
    return dt.format("D MMM YYYY");
  };

  let description = props.articleDescription.length > 145 ? props.articleDescription.substring(0, Math.min(props.articleDescription.length,145)).concat("...") : props.articleDescription;

  return <div className="articleCardTwo boxHover">
    <a
      href={props.url}
      target="_blank"
      data-interception="off"
    >
      { props.previewMediaUrl === null ?
        <div>No Image Available</div> 
        :
        <div className="twoImageContainer">
          <img 
            src={props.previewMediaUrl}
            className="articleImage"  
          />
        </div>
      }
    </a>
    <div className="articleInfo">
      <div className="articleDate">{formatDate(props.articleDate).toUpperCase()}</div>
      <div className="articleTitle">
        <a href={props.url} target="_blank" data-interceptio="off">
          {props.title}
        </a>
      </div>
    <div className="articleDescription">{description}</div>
    </div>
  </div>;
};

export const MyNewsItems = (props: IMyNewsItem) => {
  let formatDate = (date: string) => {
    let dt = moment(date);
    return dt.format("D MMM YYYY");
  };

  let displayTitle = (title:string) => {
    return title.length > 50 ? title.substring(0,47).concat("...") : title;
  };
  
  return (
    <div className="articleCardColumn boxHover">
      <a
        href={props.url}
        target="_blank"
        data-interception="off"
      >
        { props.previewMediaUrl === null ?
          <div>No Image Available</div> 
          :
          <div className="stackImageContainer">
            <img 
              src={props.previewMediaUrl}
              className="articleImage"  
            />
          </div>
        }
      </a>
      <div className="articleInfo">
        <div className="articleDate">{formatDate(props.articleDate).toUpperCase()}</div>
        <div className="articleTitle">
          <a href={props.url} target="_blank" data-interceptio="off" title={props.title}>
            {displayTitle(props.title)}
          </a>
      </div>
      </div>
    </div>
  );
};

export default class MyNews extends React.Component<IMyNewsProps, MyNewsState, {}> {
  constructor(props: IMyNewsProps){
    super(props);
    this.state = {
      newsArticles: [],
      isLoading: true,
    };
  }

  public componentDidMount() {
    this.getMyNewsData().then(data => {
      this.setState({
        newsArticles: data,
        isLoading: false,
      });
    });
  }

  private async GetUserSubscriptions() {
    sp.setup({
      sp: {
        baseUrl: this.props.siteUrl,
      }
    });

    const response = await sp.profiles.myProperties.get();
    const topicSubscriptions = response.UserProfileProperties.find(cell => cell.Key === "SubscribedTopics").Value.split("|");
    const communitiesSubscriptions = response.UserProfileProperties.find(cell => cell.Key === "SubscribedCommunities").Value.split("|");
    //const userSubscriptions = topicSubscriptions.concat("|",communitiesSubscriptions).split("|");
    let userSubscriptions: IUserSubscriptions = {
      topicSubscriptions: topicSubscriptions,
      communitySubscriptions: communitiesSubscriptions
    };
    
    return userSubscriptions;
  }

  private async getMyNewsData() {

    let spOpts = {
      headers: {
        'odata-version': '3.0',
        'Content-Type': 'application/json; odata=verbose'
      }
    };

    let userSubscriptions = await this.GetUserSubscriptions();
    let firstTopic = userSubscriptions.topicSubscriptions.toString().length;
    let topicFilters = firstTopic !== 0 ? 
      userSubscriptions.topicSubscriptions.map(item => {
        let filter = 'RefinableTopic:equals("';
        return filter.concat(item,'")');
      })
      :
      //if no topic subscriptions is blank - then do what?
      ``;

    let firstCommunity = userSubscriptions.communitySubscriptions.toString().length;
    let communityFilters = firstCommunity !== 0 ? 
      userSubscriptions.communitySubscriptions.map(item => {
        let filter = 'RefinableTeam:equals("';
        return filter.concat(item,'")');
      })
      :
      //if no community subscriptions is blank - then do what?
      ``;

    //set assigned news filter to the the current site and exclude the home page
    let currentSiteFilter =  'Path:"' + this.props.siteUrl + '" AND Path<>"' + this.props.siteUrl + '"';

    //if topics and communities are both blank
    let subscriptionFilters = topicFilters.length==0 && communityFilters.length==0 && "&";
    //if topics are blank and communities are just one
    subscriptionFilters = topicFilters.length==0 && communityFilters.length==1 ? `&refinementfilters='` + communityFilters.toString() + `'&` : subscriptionFilters;
    //if topics are blank and communities are not blank
    subscriptionFilters = topicFilters.length==0 && communityFilters.length>1 ? `&refinementfilters='or(` + communityFilters.toString() + `)'&` : subscriptionFilters;
    //if communities are blank and topics are just one
    subscriptionFilters = communityFilters.length==0 && topicFilters.length==1 ? `&refinementfilters='` + topicFilters.toString() +`'&` : subscriptionFilters;
    //if communities are blank and topics are not blank
    subscriptionFilters = communityFilters.length==0 && topicFilters.length>1 ? `&refinementfilters='or(` + topicFilters.toString() +`)'&` : subscriptionFilters;
    //if topics and communities are both not blank
    subscriptionFilters = topicFilters.length>0 && communityFilters.length>0 ? `&refinementfilters='or(` + topicFilters.toString() + `,` + communityFilters.toString() + `)'&` : subscriptionFilters;

    const myNewsArray = [];
    const selectProperties = `'Created%2cPath%2cUrl%2cTitle%2cAuthor%2cRefinableNewsType%2cTeam%2cDescription%2cFirstPublishedDate%2cPictureThumbnailURL%2cRefinableTeam%2cRefinableTeamType%2cRefinableTopic%2cSite'`;

    let searchString = `/_api/search/query?querytext='ContentType:"News Story" AND `+ currentSiteFilter +`'&trimduplicates=true&rowlimit=4&selectproperties=`+ selectProperties  + subscriptionFilters + `sortlist='Created:descending'&clienttype='ContentSearchRegular'`;

    let response = await this.props.context.spHttpClient.get(this.props.siteUrl + searchString, SPHttpClient.configurations.v1, spOpts);
    let responseJSON = await response.json();
    const myNewsArticles = responseJSON.PrimaryQueryResult.RelevantResults.Table.Rows;
    let articleIndex = 0;
    myNewsArticles.map(async article => {
      articleIndex = articleIndex + 1;
      const Article = {
        previewMediaUrl: article.Cells.find(cell => cell.Key === "PictureThumbnailURL").Value,
        title: article.Cells.find(cell => cell.Key === "Title").Value,
        articleDate: article.Cells.find(cell => cell.Key === "Created").Value,
        url: article.Cells.find(cell => cell.Key === "Path").Value,
        articleDescription:  article.Cells.find(cell => cell.Key === "Description").Value,
        articleIndex: articleIndex,
      };
      myNewsArray.push(Article);
    });

    return await Promise.all(myNewsArray);
  }

  public render(): React.ReactElement<IMyNewsProps> {
    const {
      newsArticles,
      isLoading,
    } = this.state;
    const {
      actionText,
      actionUrl,
      actionTextLeft,
      actionUrlLeft,
      title
    } = this.props;

    const headingLinkRight : ILink = {
      text: actionText,
      url: actionUrl
    };

    const headingLinkLeft : ILink = {
      text: actionTextLeft,
      url: actionUrlLeft
    };
    
    const numberOfResults = newsArticles.length;

    return (
      <div className={"myNews"}>
        <Heading heading={title} leftLink={headingLinkLeft} rightLink={headingLinkRight}/>
        <div className={"newsItemContainer"}>
          {isLoading && 
            <Spinner 
              label="Loading News..."
              size={SpinnerSize.large}
            />
          }
          { //if there are no news article returned
            !isLoading && numberOfResults==0 &&
              <div className="noNewsMessage">No news currently available.</div>
          }
          { //if there is only 1 news article returned
            !isLoading && numberOfResults==1 && newsArticles && newsArticles.map(article => {
              return <MyNewsItemOne {...article} />;
            })
          }
          { //if there are 2 news articles returned
            !isLoading && numberOfResults==2 && newsArticles && newsArticles.map(article => {
              return <MyNewsItemTwo {...article} />;
            })
          }
          { //if there more than 2 news articles returned, render first article
            !isLoading && numberOfResults>2 && newsArticles && newsArticles.map(article => {
              return article.articleIndex == 1 && <MyNewsItemTwo {...article} />;
            })
          }
            { // if there are more than 2 articles returned, return articles stacked
              !isLoading && numberOfResults>2 && 
              <div className={"articleStack"}>
                {newsArticles && newsArticles.map(article => {
                    return article.articleIndex > 1 && <MyNewsItems {...article} />;
                  })}
              </div>
            }

        </div>

      </div>
    );
  }
}
