
import { TinyColor } from '@ctrl/tinycolor';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import * as React from 'react';
import { IParam } from '../models/IParam';

export interface IPowerAppsViewer {
  appUrl: string;
  params: IParam[];
  locale: string;
  passingThemeColorsAsParams: boolean;
  themeColorsParamPrefix: string;
  theme: IReadonlyTheme;
  width?: string | number;
  height?: string | number;
  showBorder?: boolean;
}

export function validateAppUrl(url: string): boolean {
  const appUrl = new URL(url);
  const hostNameAllowList = [
    "apps.powerapps.com",
    "apps.gov.powerapps.us",
  ];
  return hostNameAllowList.indexOf(appUrl.hostname.toLowerCase()) !== -1 ? true : false;
}

function generateIFrameUrl(props: IPowerAppsViewer): string {

  const checkAppUrl = props.appUrl && props.appUrl !== '' && validateAppUrl(props.appUrl);

  if (checkAppUrl) {
    try {
      const url = new URL(props.appUrl);

      const screenColorRgb = new TinyColor(props.theme.semanticColors.bodyBackground).toRgb();
      url.searchParams.set('screenColor', `rgba(${screenColorRgb.r},${screenColorRgb.g},${screenColorRgb.b},1)`);
      url.searchParams.set('source', 'Apvee-PowerApps-iFrame');
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
}

export default function PowerAppsViewer(props: IPowerAppsViewer): JSX.Element {

  const appUrl = generateIFrameUrl(props);

  return (
    <iframe
      name='Apvee-PowerApps-iFrame'
      src={appUrl}
      scrolling="no"
      allowFullScreen={true}
      allow="geolocation *; microphone *; camera *; fullscreen *;"
      sandbox="allow-popups allow-popups-to-escape-sandbox allow-same-origin allow-scripts allow-forms allow-orientation-lock allow-downloads"
      width={props.width}
      height={props.height}
      style={{ boxSizing: "border-box", border: props.showBorder ? `1px solid ${props.theme.semanticColors.bodyFrameDivider}` : "none" }}
      frameBorder={0}
      aria-hidden="true"
    />
  );
}
