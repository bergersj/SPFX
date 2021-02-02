export interface ILink {
  text: string;
  url: string;
  openInNewTab?: boolean;
}

export interface IQuickLinksProps {
  heading: string;
  items: ILink[];
}
