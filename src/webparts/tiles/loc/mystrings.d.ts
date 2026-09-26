declare interface ITilesWebPartStrings {
  PropertyPaneDescription: string;
  DataSourceGroupName: string;
  DisplayGroupName: string;
  BehaviorGroupName: string;
  ListTitleFieldLabel: string;
  ItemLimitFieldLabel: string;
  EnableSearchFieldLabel: string;
  EnableFilterFieldLabel: string;
  LoadingLabel: string;
  EmptyMessage: string;
  ConfigErrorMessage: string;
  UnauthorizedMessage: string;
}

declare module 'TilesWebPartStrings' {
  const strings: ITilesWebPartStrings;
  export = strings;
}
