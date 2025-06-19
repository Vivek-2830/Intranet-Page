declare interface ISharePointDirectoryWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'SharePointDirectoryWebPartStrings' {
  const strings: ISharePointDirectoryWebPartStrings;
  export = strings;
}
