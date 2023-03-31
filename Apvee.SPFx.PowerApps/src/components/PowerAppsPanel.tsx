import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import PowerAppsViewer from './PowerAppsViewer';
import { IParam } from '../models/IParam';

export interface IPowerAppsPanelProps {
    appWebLink: string;
    panelTitle: string;
    width: "small" | "medium" | "large" | "xlarge" | "full";
    params: IParam[];
    locale: string;
    passingThemeColorsAsParams: boolean;
    themeColorsParamPrefix: string;
    showBorder?: boolean;
    theme: IReadonlyTheme;
    onDismiss: () => void;
}

const getPanelWidth = (width: IPowerAppsPanelProps["width"]): PanelType => {
    switch (width) {
        case "small":
            return PanelType.smallFixedFar;
        case "medium":
            return PanelType.medium;
        case "large":
            return PanelType.large;
        case "xlarge":
            return PanelType.extraLarge;
        case "full":
            return PanelType.smallFluid;
    }
};

export default function PowerAppsPanel(props: IPowerAppsPanelProps): JSX.Element {
    return (
        <Panel
            styles={{
                scrollableContent: {
                    display: "flex",
                    flexDirection: "row",
                    flexWrap: "nowrap",
                    justifyContent: "flex-start",
                    alignItems: "stretch",
                },
                content: {
                    width: "100%"
                }
            }}
            isFooterAtBottom={true}
            isOpen={true}
            type={getPanelWidth(props.width)}
            onDismiss={() => props.onDismiss()}
            headerText={props.panelTitle}
            closeButtonAriaLabel="Close" >
            <PowerAppsViewer
                appUrl={props.appWebLink}
                params={props.params}
                locale={props.locale}
                passingThemeColorsAsParams={false}
                themeColorsParamPrefix={props.themeColorsParamPrefix}
                showBorder={props.showBorder}
                theme={props.theme}
                height='100%'
                width='100%' />
        </Panel>
    );
}