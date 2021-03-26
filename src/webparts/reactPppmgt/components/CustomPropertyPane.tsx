import { GroupType, Login, PeoplePicker, Person, TeamsChannelPicker } from '@microsoft/mgt-react';
import * as React from 'react';

import { ICustomPropertyPaneProps } from './ICustomPropertyPaneProps';

import { Provider, teamsTheme, teamsDarkTheme, teamsHighContrastTheme, FormDatepicker, FormRadioGroup } from '@fluentui/react-northstar';
import { PropertyPanePortal } from '../../../PPP/PropertyPanePortal';
import { Providers } from '@microsoft/mgt';

export const CustomPropertyPane: React.FunctionComponent<ICustomPropertyPaneProps> = (props) => {

    // Function to update Web Part properties and re-render the web Part
    function updateWPProperty(p, v) {
        props.propertyBag[p] = v;
        props.renderWP();
    }

    // Teams themes
    let currentTheme;

    switch (props.propertyBag["northstarRadioGroup"]) {
        case "Light": currentTheme = teamsTheme; break;
        case "Dark": currentTheme = teamsDarkTheme; break;
        case "Contrast": currentTheme = teamsHighContrastTheme; break;
        default: currentTheme = teamsTheme;
    }

    return (
        <>
            <Provider theme={currentTheme}>
                <Login
                    data-Property="mgtPerson"
                    loginCompleted={(e) => {
                        console.log("login completed");
                        Providers.globalProvider.graph.client.api('me').get()
                            .then(getMe => console.log(getMe));
                    }}
                    logoutCompleted={(e) => { console.log("logout completed"); }}
                />
                <PropertyPanePortal propertyPaneHosts={props.propertyPaneHosts}>
                    <PeoplePicker
                        data-Property="mgtPeoplePicker"
                        selectionMode="single"
                        defaultSelectedUserIds={[props.propertyBag.mgtPeoplePicker]}
                        selectionChanged={(e: any) => {
                            console.log(e.detail);
                            let users = [];
                            e.detail.forEach(dtl => users.push(dtl.userPrincipalName));
                            updateWPProperty("mgtPeoplePicker", users[0]);
                        }}
                    />
                    <PeoplePicker
                        data-Property="mgtGroupPicker"
                        selectionMode="single"
                        groupType={GroupType.unified}
                        defaultSelectedUserIds={[props.propertyBag.mgtPeoplePicker]}
                        selectionChanged={(e: any) => {
                            console.log(e.detail);
                            let users = [];
                            e.detail.forEach(dtl => users.push(dtl.userPrincipalName));
                            updateWPProperty("mgtGroupPicker", users[0]);
                        }}
                    />
                    {/* <TeamsChannelPicker
                    data-Property="mgtTeamsChannelPicker"
                    defaultValue={[props.propertyBag.mgtTeamsChannelPicker]}

                    selectionChanged={(e: any) => {
                        let slctns = [];
                        console.log(e);
                        e.detail.forEach(dtl => slctns.push(dtl.channel));
                        updateWPProperty("mgtTeamsChannelPicker", slctns[0]);
                    }}
                /> */}
                </PropertyPanePortal>
            </Provider>
        </>
    );
};