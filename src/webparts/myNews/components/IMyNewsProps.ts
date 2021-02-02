export interface IMyNewsItem {
  previewMediaUrl: string;
  title: string;
  articleDate: string;
  url: string;
  articleDescription: string;
  articleIndex?: number;
}

export interface IMyNewsProps {
  title: string;
  siteUrl: string;
  actionText: string;
  actionUrl: string;
  actionTextLeft: string;
  actionUrlLeft: string;
  context: any;
  siteID:any;
  //newsArticles: IMyNewsItem[];
  // spHttpClient: SPHttpClient;
}
