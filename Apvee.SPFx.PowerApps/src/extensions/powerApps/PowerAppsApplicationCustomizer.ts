//import { Log } from '@microsoft/sp-core-library';
import {
  BaseApplicationCustomizer, PlaceholderContent, PlaceholderName
} from '@microsoft/sp-application-base';
import * as React from 'react';
//import { Dialog } from '@microsoft/sp-dialog';
import * as ReactDOM from 'react-dom';
import PowerAppsPanel, { IPowerAppsPanelProps } from './components/PowerAppsPanel';

//import * as strings from 'PowerAppsApplicationCustomizerStrings';
//const LOG_SOURCE: string = 'PowerAppsApplicationCustomizer';

export interface IPowerAppsApplicationCustomizerPostMessage {
  type: "OpenPowerAppsViewer" | "ClosePowerAppsViewer";
  powerAppUrl?: string;
  panelTitle?: string;
  panelSize?: string;
}

export interface IPowerAppsApplicationCustomizerProperties { }

export default class PowerAppsApplicationCustomizer extends BaseApplicationCustomizer<IPowerAppsApplicationCustomizerProperties> {

  private topPlaceholder: PlaceholderContent | undefined;

  public onInit(): Promise<void> {
    this.context.placeholderProvider.changedEvent.add(this, this.initPlaceholder);
    return Promise.resolve();
  }

  private initPlaceholder(): void {
    if (!this.topPlaceholder) {
      this.topPlaceholder = this.context.placeholderProvider.tryCreateContent(PlaceholderName.Top, {
        onDispose: () => { ReactDOM.unmountComponentAtNode(this.topPlaceholder.domElement) }
      });

      if (!this.topPlaceholder) {
        console.error("The expected placeholder (Top) was not found.");
        return;
      }

      window.addEventListener('message', (event) => {
        const message: IPowerAppsApplicationCustomizerPostMessage = event.data;

        if (message && (message.type === 'OpenPowerAppsViewer' || message.type === 'ClosePowerAppsViewer')) {
          this.interceptRequest(message);
        }
      });
    }
  }

  private interceptRequest(message: IPowerAppsApplicationCustomizerPostMessage): void {
    if (message.type === 'OpenPowerAppsViewer') {
      //this.topPlaceholder.domElement.innerHTML = JSON.stringify(message);
      const element: React.ReactElement<IPowerAppsPanelProps> = React.createElement(
        PowerAppsPanel, {}
      );

      ReactDOM.render(element, this.topPlaceholder.domElement);
    }

    if (message.type === 'ClosePowerAppsViewer') {
      //this.topPlaceholder.domElement.innerHTML = '';
      ReactDOM.unmountComponentAtNode(this.topPlaceholder.domElement)
    }

    // Dialog.alert(`Hello ${message}`).catch(() => {
    //   /* handle error */
    // });
  }
}
