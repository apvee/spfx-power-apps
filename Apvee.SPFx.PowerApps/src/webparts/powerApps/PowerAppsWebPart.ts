import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneDropdown,
  PropertyPaneTextField,
  PropertyPaneToggle,
  PropertyPaneDynamicFieldSet,
  PropertyPaneDynamicField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart, IWebPartPropertiesMetadata } from '@microsoft/sp-webpart-base';
import { IReadonlyTheme } from '@microsoft/sp-component-base';

import * as strings from 'PowerAppsWebPartStrings';
import { IParams } from '../../models/IParams';
import { AspectRatio } from '../../models/AspectRatio';
import PowerApps, { IPowerAppsProps } from './components/PowerApps';
import { PropertyFieldCollectionData, CustomCollectionFieldType } from '@pnp/spfx-property-controls/lib/PropertyFieldCollectionData';
import { DynamicProperty } from '@microsoft/sp-component-base';
import { PropertyFieldSpinButton } from '@pnp/spfx-property-controls/lib/PropertyFieldSpinButton';

export interface IPowerAppsWebPartProps {
  title: string;
  appWebLink: string;
  params: IParams[];
  passingThemeColorsAsParams: boolean;
  showBorder: boolean;
  useCustomHeight: boolean;
  customHeight: number;
  aspectRatio: AspectRatio;
  useDynamicProp: boolean;
  dynamicPropName: string;
  dynamicProp: DynamicProperty<string>;
}

export default class PowerAppsWebPart extends BaseClientSideWebPart<IPowerAppsWebPartProps> {

  private currentTheme: IReadonlyTheme | undefined;

  public render(): void {

    const dynamicProp: string | undefined = this.properties.dynamicProp?.tryGetValue();
    const locale: string = this.context.pageContext.cultureInfo.currentCultureName;

    const element: React.ReactElement<IPowerAppsProps> = React.createElement(
      PowerApps,
      {
        title: this.properties.title,
        appWebLink: this.properties.appWebLink,
        params: this.properties.params,
        locale: locale,
        passingThemeColorsAsParams: this.properties.passingThemeColorsAsParams,

        useDynamicProp: this.properties.useDynamicProp,
        dynamicPropName: this.properties.dynamicPropName,
        dynamicProp: dynamicProp,

        theme: this.currentTheme,
        showBorder: this.properties.showBorder,

        useCustomHeight: this.properties.useCustomHeight,
        customHeight: this.properties.customHeight,
        aspectRatio: this.properties.aspectRatio,

        displayMode: this.displayMode,
        updateTitle: (value: string) => { this.properties.title = value },
        openPropertyPane: this.context.propertyPane.open
      }
    );

    ReactDom.render(element, this.domElement);

    console.log(this.currentTheme);
  }

  protected onThemeChanged(currentTheme: IReadonlyTheme | undefined): void {
    this.currentTheme = currentTheme;
  }

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
  }

  protected get propertiesMetadata(): IWebPartPropertiesMetadata {
    return {
      // Specify the web part properties data type to allow the address
      // information to be serialized by the SharePoint Framework.
      'dynamicProp': {
        dynamicPropertyType: 'string'
      }
    };
  }

  protected getPropertyPaneConfiguration(): IPropertyPaneConfiguration {
    return {
      pages: [
        {
          header: {
            description: strings.PropertyPaneDescription
          },
          groups: [
            {
              groupName: strings.BasicGroupName,
              groupFields: [
                PropertyPaneTextField('appWebLink', {
                  label: strings.AppWebLinkLabel
                }),
                PropertyPaneToggle('useCustomHeight', {
                  label: "useCustomHeight"
                }),
                PropertyFieldSpinButton('customHeight', {
                  label: 'customHeight',
                  disabled: !this.properties.useCustomHeight,
                  initialValue: this.properties.customHeight,
                  onPropertyChange: this.onPropertyPaneFieldChanged,
                  properties: this.properties,
                  suffix: 'px',
                  min: 0,
                  key: 'customHeightFieldId'
                }),
                PropertyPaneDropdown('aspectRatio', {
                  label: strings.AspectRatioLabel,
                  disabled: this.properties.useCustomHeight,
                  options: [
                    { key: '16:9', text: '16:9' },
                    { key: '3:2', text: '3:2' },
                    { key: '16:10', text: '16:10' },
                    { key: '4:3', text: '4:3' },
                    { key: '1:1', text: '1:1' },
                    { key: '3:4', text: '3:4' },
                    { key: '10:16', text: '10:16' },
                    { key: '2:3', text: '2:3' },
                    { key: '9:16', text: '9:16' }
                  ]
                }),
                PropertyPaneToggle('showBorder', {
                  label: strings.ShowBorderLabel
                })
              ]
            },
            {
              groupName: strings.ParametersGroupName,
              groupFields: [
                PropertyPaneToggle('passingThemeColorsAsParams', {
                  label: strings.PassingThemeColorsAsParamsLabel,
                }),
                PropertyFieldCollectionData("params", {
                  key: "params",
                  label: strings.ParamsLabel,
                  panelHeader: strings.ParamsPanelHeader,
                  manageBtnLabel: strings.ParamsManageBtnLabel,
                  value: this.properties.params,
                  fields: [
                    {
                      id: "name",
                      title: strings.ParamsNameLabel,
                      type: CustomCollectionFieldType.string,
                      required: true
                    },
                    {
                      id: "value",
                      title: strings.ParamsValueLabel,
                      type: CustomCollectionFieldType.string
                    }
                  ],
                  disabled: false,
                  enableSorting: true
                }),
                PropertyPaneToggle('useDynamicProp', {
                  label: strings.UseDynamicPropLabel,
                  checked: this.properties.useDynamicProp === true
                }),
                this.properties.useDynamicProp === true && PropertyPaneTextField('dynamicPropName', {
                  label: strings.DynamicPropNameLabel,
                  value: this.properties.dynamicPropName
                }),
                this.properties.useDynamicProp === true && PropertyPaneDynamicFieldSet({
                  label: strings.SelectDynamicSourceLabel,
                  fields: [
                    PropertyPaneDynamicField('dynamicProp', {
                      label: strings.SelectDynamicPropFieldLabel
                    })
                  ]
                })
              ]
            },
          ]
        }
      ]
    };
  }
}
