
import { TinyColor } from '@ctrl/tinycolor';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as React from 'react';
import { IParams } from '../models/IParams';

export interface IPowerAppsProps {
  appWebLink: string;
  params: IParams[];
  locale: string;
  passingThemeColorsAsParams: boolean;
  themeColorsParamPrefix: string;
  theme: IReadonlyTheme;
}

const generateUrl = (props: IPowerAppsProps): string => {
  if (props.appWebLink) {
    try {
      const url = new URL(props.appWebLink);

      const screenColorRgb = new TinyColor(props.theme.semanticColors.bodyBackground).toRgb();
      url.searchParams.set('screenColor', `rgba(${screenColorRgb.r},${screenColorRgb.g},${screenColorRgb.b},1)`);
      url.searchParams.set('source', 'Apvee-PowerAppsAplicationCustomizer');
      url.searchParams.set('locale', props.locale);

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
          url.searchParams.set(`${props.themeColorsParamPrefix}${paletteKey}`, (props.theme.palette as any)[paletteKey]);
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

export default function PowerAppsViewer(props: IPowerAppsProps): JSX.Element {

  const appUrl = generateUrl(props);
  const isConfigured = checkMandatoryProps(props);

  return (
    isConfigured ?
      <iframe
        src={appUrl}
        aria-hidden="true"
        allow="geolocation *; microphone *; camera *; fullscreen *;"
        sandbox="allow-popups allow-popups-to-escape-sandbox allow-same-origin allow-scripts allow-forms allow-orientation-lock allow-downloads"
        frameBorder={0}
      />
      : <></>
  );
}
