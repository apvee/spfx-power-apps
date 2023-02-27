import { Panel, PanelType } from 'office-ui-fabric-react/lib/Panel';
import * as React from 'react';

export interface IPowerAppsPanelProps { }
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
            type={PanelType.smallFixedFar}
            onDismiss={() => { window.postMessage({ "type": "ClosePowerAppsViewer" }) }}
            headerText="Header"
            closeButtonAriaLabel="Close" >
            <div
                style={{
                    background: 'green',
                    height: '100%',
                    width: '100%'
                }} >
                Stuff
            </div>
        </Panel>
    );
}