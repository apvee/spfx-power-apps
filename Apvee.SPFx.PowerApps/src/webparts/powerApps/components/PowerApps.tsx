
import { TinyColor } from '@ctrl/tinycolor';
import { AspectRatio } from '@mantine/core';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import { DisplayMode } from '@microsoft/sp-core-library';
import { Placeholder } from "@pnp/spfx-controls-react/lib/Placeholder";
import { WebPartTitle } from "@pnp/spfx-controls-react/lib/WebPartTitle";
import { Stack } from 'office-ui-fabric-react';
import * as strings from 'PowerAppsWebPartStrings';
import * as React from 'react';
import { AspectRatio as AspectRatioType } from '../../../models/AspectRatio';
import { IParams } from '../../../models/IParams';

export interface IPowerAppsProps {
  title: string;
  appWebLink: string;
  params: IParams[];
  locale: string;
  passingThemeColorsAsParams: boolean;

  useDynamicProp: boolean;
  dynamicPropName: string,
  dynamicProp: string,

  theme: IReadonlyTheme;
  showBorder: boolean;
  aspectRatio: AspectRatioType;

  displayMode: DisplayMode;
  updateTitle: (value: string) => void;
  openPropertyPane: () => void;
}

const generateUrl = (props: IPowerAppsProps): string => {
  if (props.appWebLink) {
    try {
      const url = new URL(props.appWebLink);

      const screenColorRgb = new TinyColor(props.theme.semanticColors.bodyBackground).toRgb();
      url.searchParams.set('screenColor', `rgba(${screenColorRgb.r},${screenColorRgb.g},${screenColorRgb.b},1)`);
      url.searchParams.set('source', 'Apvee-PowerAppsWebPart');
      url.searchParams.set('locale', props.locale);

      if (props.useDynamicProp) {
        url.searchParams.set(props.dynamicPropName, props.dynamicProp);
      }

      if (props.params) {
        props.params.forEach(param => {
          if (param.value.trim() !== '') {
            url.searchParams.set(param.name, param.value);
          }
        });
      }

      if (props.passingThemeColorsAsParams) {
        Object.keys(props.theme.palette).forEach((paletteKey: string) => {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          url.searchParams.set(`theme-${paletteKey}`, (props.theme.palette as any)[paletteKey]);
        });
      }

      return url.toString();
    } catch (e) {
      return null;
    }
  } else {
    return null;
  }
};

const generateBorder = (showBorder: boolean, theme: IReadonlyTheme): React.CSSProperties => {
  if (showBorder)
    return {
      border: `1px solid ${theme.semanticColors.bodyFrameDivider}`
    }
  else
    return {};
}

const checkMandatoryProps = (props: IPowerAppsProps): boolean => {
  let result = props.appWebLink && props.appWebLink !== '';

  if (result) {
    try {
      const appUrl = new URL(props.appWebLink);
      result = appUrl.hostname.toLowerCase() === "apps.powerapps.com" || appUrl.hostname.toLowerCase() === "apps.gov.powerapps.us";
    } catch (e) {
      result = false;
    }
  }

  return result;
}

export default function PowerApps(props: IPowerAppsProps): JSX.Element {

  const [ratio, setRatio] = React.useState(16 / 9);
  const appUrl = generateUrl(props);
  const isConfigured = checkMandatoryProps(props);

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

  return (
    <Stack>
      <WebPartTitle
        displayMode={props.displayMode}
        title={props.title}
        updateProperty={props.updateTitle}
        themeVariant={props.theme} />
      {!isConfigured &&
        <Placeholder
          iconName='PowerApps'
          iconText={strings.PlaceholderIconText}
          description={strings.PlaceholderDescription}
          buttonLabel={strings.PlaceholderButtonLabel}
          onConfigure={props.openPropertyPane}
          hideButton={props.displayMode === DisplayMode.Read}
          theme={props.theme} />
      }
      {isConfigured &&
        <AspectRatio ratio={ratio}>
          <iframe
            src={appUrl}
            aria-hidden="true"
            allow="geolocation *; microphone *; camera *; fullscreen *;"
            sandbox="allow-popups allow-popups-to-escape-sandbox allow-same-origin allow-scripts allow-forms allow-orientation-lock allow-downloads"
            frameBorder={0}
            style={generateBorder(props.showBorder, props.theme)}
          />
        </AspectRatio>
      }
    </Stack>
  );
}
