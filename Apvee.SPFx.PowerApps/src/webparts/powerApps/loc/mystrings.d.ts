declare interface IPowerAppsWebPartStrings {
  PropertyPaneDescription: string;
  BasicGroupName: string;
  AppWebLinkLabel: string;
  ShowBorderLabel: string;
  AspectRatioLabel: string;
  ParametersGroupName: string;
  PassingThemeColorsAsParamsLabel: string;
  ThemeColorsParamPrefixLabel: string;
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
  UseCustomHeightLabel: string;
  CustomHeightLabel: string;

  ShowAsPanelLabel: string;
  ButtonOpenPanelTextLabel: string;
  ButtonOpenPanelPositionLabel: string;
  PanelTitleLabel: string;
  PanelWidthLabel: string;

  StartLabel: string;
  CenterLabel: string;
  EndLabel: string;

  SmallLabel: string;
  MediumLabel: string;
  LargeLabel: string;
  XlargeLabel: string;
  FullLabel: string;
}

declare module 'PowerAppsWebPartStrings' {
  const strings: IPowerAppsWebPartStrings;
  export = strings;
}
