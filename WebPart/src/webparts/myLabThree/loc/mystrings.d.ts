declare interface IMyLabThreeWebPartStrings {
  PropertyPaneDescription: string;
  DisplayGroupName: string;
  DescriptionFieldLabel: string;
  AlertMessage: string;
  LinkAddress: string;
  LinkTextDisplay: string;
  LocalMessage: string;
  OnlineMessage: string;
  ButtonLocaleName: string;
  LabelLocaleText: string;
  WelcomeMessage: string;
  IntroductionMessage: string;
  LearnLocaleName: string;
  LearnMoreLinkAddress: string;
}

declare module 'MyLabThreeWebPartStrings' {
  const strings: IMyLabThreeWebPartStrings;
  export = strings;
}
