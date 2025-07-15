declare interface IHelloWorldWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  DescriptionFieldLabel: string;
  AppLocalEnvironmentSharePoint: string;
  AppLocalEnvironmentTeams: string;
  AppLocalEnvironmentOffice: string;
  AppLocalEnvironmentOutlook: string;
  AppSharePointEnvironment: string;
  AppTeamsTabEnvironment: string;
  AppOfficeEnvironment: string;
  AppOutlookEnvironment: string;
  UnknownEnvironment: string;
  //Added by Nilangi
  /*BotUrlFieldLabel: string;
  BotNameFieldLabel: string;
  ButtonLabelFieldLabel: string;
  CustomScopeFieldLabel: string;
  ClientIdFieldLabel: string;
  AuthorityFieldLabel: string;
  GreetFieldLabel: string;*/
}

declare module 'HelloWorldWebPartStrings' {
  const strings: IHelloWorldWebPartStrings;
  export = strings;
}
