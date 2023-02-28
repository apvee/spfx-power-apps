import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';
import { IReadonlyTheme } from '@microsoft/sp-component-base';
import PowerAppsViewer from './PowerAppsViewer';
import { IParam } from '../models/IParam';

export interface IPowerAppsPanelProps {
    panelTitle: string;
    appWebLink: string;
    params: IParam[];
    locale: string;
    passingThemeColorsAsParams: boolean;
    themeColorsParamPrefix: string;
    theme: IReadonlyTheme;
    onDismiss: () => void;
}
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
            type={PanelType.medium}
            onDismiss={() => props.onDismiss()}
            headerText={props.panelTitle}
            closeButtonAriaLabel="Close" >
            <PowerAppsViewer
                appWebLink={props.appWebLink}
                params={props.params}
                locale={props.locale}
                passingThemeColorsAsParams={false}
                themeColorsParamPrefix={props.themeColorsParamPrefix}
                theme={props.theme}
                height='100%'
                width='100%' />
        </Panel>
    );
}