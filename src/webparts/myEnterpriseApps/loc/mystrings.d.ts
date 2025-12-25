declare interface IMyEnterpriseAppsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  SortOrderFieldLabel: string;
  SortOrderFieldDescription: string;
  ShowHiddenAppsLabel: string;
  IconSizeFieldLabel: string;
  IconSizeSmall: string;
  IconSizeNormal: string;
  IconSizeLarge: string;
  IconSizeHuge: string;
  DefaultTitle: string;
  AllAppsLabel: string;
  NoAppsFound: string;
  ErrorLoading: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
}

declare module 'MyEnterpriseAppsWebPartStrings' {
  const strings: IMyEnterpriseAppsWebPartStrings;
  export = strings;
}
