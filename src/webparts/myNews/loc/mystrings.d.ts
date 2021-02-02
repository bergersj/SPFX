declare interface IMyNewsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
}

declare module 'MyNewsWebPartStrings' {
  const strings: IMyNewsWebPartStrings;
  export = strings;
}
