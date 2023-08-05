
import { AspectRatio } from '@mantine/core';
import { useToggle } from '@mantine/hooks';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { DefaultButton, Stack, ThemeProvider } from 'office-ui-fabric-react';
import * as strings from 'PowerAppsWebPartStrings';
import * as React from 'react';
import PowerAppsPanel from '../../../components/PowerAppsPanel';
import PowerAppsViewer, { validateAppUrl } from '../../../components/PowerAppsViewer';
import { AspectRatio as AspectRatioType } from '../../../models/AspectRatio';
import { IParam } from '../../../models/IParam';

export interface IPowerAppsProps {
  title: string;
  appWebLink: string;
  params: IParam[];
  locale: string;
  passingThemeColorsAsParams: boolean;
  themeColorsParamPrefix: string;

  useDynamicProp: boolean;
  dynamicPropName: string,
  dynamicProp: string,

  theme: IReadonlyTheme;
  showBorder: boolean;

  useCustomHeight: boolean;
  customHeight: number;
  aspectRatio: AspectRatioType;

  showAsPanel: boolean;
  buttonOpenPanelText: string;
  buttonOpenPanelPosition: "start" | "center" | "end";
  panelTitle: string;
  panelWidth: "small" | "medium" | "large" | "xlarge" | "full";

  displayMode: DisplayMode;
  updateTitle: (value: string) => void;
  openPropertyPane: () => void;
}

export default function PowerApps(props: IPowerAppsProps): JSX.Element {
  const [ratio, setRatio] = React.useState(16 / 9);
  const isConfigured = validateAppUrl(props.appWebLink);
  const params = React.useState<IParam[]>([]);
  const [isPanelOpen, setIsPanelOpen] = useToggle();

  React.useEffect(() => {
    switch (props.aspectRatio) {
      case '16:9':
        setRatio(16 / 9);
        break;
      case '3:2':
        setRatio(3 / 2);
        break;
      case '16:10':
        setRatio(16 / 10);
        break;
      case '4:3':
        setRatio(4 / 3);
        break;
      case '1:1':
        setRatio(1 / 1);
        break;
      case '3:4':
        setRatio(3 / 4);
        break;
      case '10:16':
        setRatio(10 / 16);
        break;
      case '2:3':
        setRatio(2 / 3);
        break;
      case '9:16':
        setRatio(9 / 16);
        break;
      default:
        setRatio(16 / 9);
        break;
    }
  }, [props.aspectRatio]);

  React.useEffect(() => {
    params.push(...[props.params]);
    if (props.useDynamicProp) {
      props.params.push({ name: props.dynamicPropName, value: props.dynamicProp });
    }
  }, [props.params, props.useDynamicProp, props.dynamicProp]);

  let elementToRender: JSX.Element;
  if (!isConfigured) {
    elementToRender =
      <Placeholder
        iconName='PowerApps'
        iconText={strings.PlaceholderIconText}
        description={strings.PlaceholderDescription}
        buttonLabel={strings.PlaceholderButtonLabel}
        onConfigure={props.openPropertyPane}
        hideButton={props.displayMode === DisplayMode.Read}
        theme={props.theme} />
  } else {

    if (props.showAsPanel) {
      elementToRender =
        <>
          <Stack horizontalAlign={props.buttonOpenPanelPosition}>
            <DefaultButton text={props.buttonOpenPanelText} onClick={() => setIsPanelOpen(true)} />
          </Stack>
          {isPanelOpen && <PowerAppsPanel
            panelTitle={props.panelTitle}
            appWebLink={props.appWebLink}
            width={props.panelWidth}
            params={props.params}
            locale={props.locale}
            passingThemeColorsAsParams={props.passingThemeColorsAsParams}
            themeColorsParamPrefix={props.themeColorsParamPrefix}
            showBorder={props.showBorder}
            theme={props.theme}
            onDismiss={() => setIsPanelOpen(false)} />}
        </>;
    } else {
      if (props.useCustomHeight === false) {
        elementToRender =
          <AspectRatio ratio={ratio}>
            <PowerAppsViewer
              appUrl={props.appWebLink}
              params={props.params}
              locale={props.locale}
              passingThemeColorsAsParams={props.passingThemeColorsAsParams}
              themeColorsParamPrefix={props.themeColorsParamPrefix}
              showBorder={props.showBorder}
              theme={props.theme} />
          </AspectRatio>;
      } else {
        elementToRender =
          <PowerAppsViewer
            appUrl={props.appWebLink}
            params={props.params}
            locale={props.locale}
            passingThemeColorsAsParams={props.passingThemeColorsAsParams}
            themeColorsParamPrefix={props.themeColorsParamPrefix}
            showBorder={props.showBorder}
            theme={props.theme}
            width="100%"
            height={`${props.customHeight}px`}
          />;
      }
    }
  }

  return (
    <ThemeProvider theme={props.theme}>
      <Stack>
        <WebPartTitle
          displayMode={props.displayMode}
          title={props.title}
          updateProperty={props.updateTitle}
          themeVariant={props.theme} />
        {elementToRender}
      </Stack>
    </ThemeProvider>
  );
}
