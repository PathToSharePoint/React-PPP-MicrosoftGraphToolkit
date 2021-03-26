import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneHorizontalRule,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'ReactPPPMGTWebPartStrings';
import ReactPPPMGT from './components/ReactPPPMGT';
import { IReactPPPMGTProps } from './components/IReactPPPMGTProps';
import { CustomPropertyPane } from './components/CustomPropertyPane';
import { IPropertyPaneHostsProps, PropertyPaneHostsFactory } from '../../PPP/PropertyPaneHostsStore';
import { PropertyPaneHost } from '../../PPP/PropertyPaneHost';

import { Providers, SharePointProvider } from '@microsoft/mgt';

export interface IReactPPPMGTWebPartProps {
  description: string;
  mgtPerson: string;
  mgtPeoplePicker: string;
  mgtGroupPicker: string;
}

export default class ReactPPPMGTWebPart extends BaseClientSideWebPart<IReactPPPMGTWebPartProps> {
  
  public onInit(): Promise<void> {

    return super.onInit().then(_ => {
  
      // other init code may be present
      Providers.globalProvider = new SharePointProvider(this.context);
      
    });
  }

  public render(): void {
    ReactDom.render(
      <>
        {/* Web Part content */}
        < ReactPPPMGT {...this.properties} />
        {/* Property Pane custom controls */}
        < CustomPropertyPane
          propertyBag={this.properties}
          renderWP={this.render.bind(this)}
          propertyPaneHosts={this.propertyPaneHosts}
        />
      </>,
      this.domElement);
  }

  // Store for managing the Property Pane hosts
  public propertyPaneHosts: IPropertyPaneHostsProps = PropertyPaneHostsFactory();

  protected onDispose(): void {
    ReactDom.unmountComponentAtNode(this.domElement);
  }

  protected get dataVersion(): Version {
    return Version.parse('1.0');
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
                // PropertyPaneHost is a generic control that hosts the actual control
                PropertyPaneHorizontalRule(),
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneHorizontalRule(),
                PropertyPaneHost('mgtPerson', this.propertyPaneHosts),
                PropertyPaneHorizontalRule(),
                PropertyPaneHost('mgtPeoplePicker', this.propertyPaneHosts),
                PropertyPaneHorizontalRule(),
                PropertyPaneHost('mgtGroupPicker', this.propertyPaneHosts),
                PropertyPaneHorizontalRule(),
          ]
        }
      ]
    }
      ]
  };
}
}
