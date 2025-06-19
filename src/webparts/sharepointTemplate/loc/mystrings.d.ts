declare interface ISharepointTemplateWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
}

declare module 'SharepointTemplateWebPartStrings' {
  const strings: ISharepointTemplateWebPartStrings;
  export = strings;
}
