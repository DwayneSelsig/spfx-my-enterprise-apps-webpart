declare interface IMyEnterpriseAppsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  TitleFieldLabel: string;
  SortOrderFieldLabel: string;
  SortOrderFieldDescription: string;
  ShowHiddenAppsLabel: string;
  ShowDefaultAppsLabel: string;
  LayoutPresetFieldLabel: string;
  LayoutPresetSmall: string;
  LayoutPresetNormal: string;
  LayoutPresetLarge: string;
  LayoutPresetHuge: string;
  LayoutPresetCustom: string;
  IconSizeFieldLabel: string;
  TextSizeFieldLabel: string;
  AppSpacingFieldLabel: string;
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
