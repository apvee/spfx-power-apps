declare interface IPowerAppsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  AppWebLinkLabel: string;
  ShowBorderLabel: string;
  AspectRatioLabel: string;
  ParametersGroupName: string;
  PassingThemeColorsAsParamsLabel: string;
  ParamsLabel: string;
  ParamsPanelHeader: string;
  ParamsManageBtnLabel: string;
  ParamsNameLabel: string;
  ParamsValueLabel: string;
  UseDynamicPropLabel: string;
  DynamicPropNameLabel: string;
  SelectDynamicSourceLabel: string;
  SelectDynamicPropFieldLabel: string;
  PlaceholderIconText: string;
  PlaceholderDescription: string;
  PlaceholderButtonLabel: string;
}

declare module 'PowerAppsWebPartStrings' {
  const strings: IPowerAppsWebPartStrings;
  export = strings;
}
