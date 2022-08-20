import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  IPropertyPaneConfiguration,
  PropertyPaneTextField
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';

import * as strings from 'NyRpgAdminSearchWebPartStrings';
import NyRpgAdminSearch from './components/NyRpgAdminSearch';
import { INyRpgAdminSearchProps } from './components/INyRpgAdminSearchProps';

import { sp } from "@pnp/sp";

export interface INyRpgAdminSearchWebPartProps {
  description: string;
  AdminGroupID: number;
}

export default class NyRpgAdminSearchWebPart extends BaseClientSideWebPart<INyRpgAdminSearchWebPartProps> {

  public render(): void {
    const element: React.ReactElement<INyRpgAdminSearchProps> = React.createElement(
      NyRpgAdminSearch,
      {
        context: this.context,
        spHttpClient: this.context.spHttpClient,  
        AdminGroupID: this.properties.AdminGroupID,
      }
    );

    ReactDom.render(element, this.domElement);
  }

  public onInit(): Promise<void> {
    return super.onInit().then(_ => {
    sp.setup({
    spfxContext: this.context,
    });
    });
    }

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
                PropertyPaneTextField('description', {
                  label: strings.DescriptionFieldLabel
                }),
                PropertyPaneTextField('AdminGroupID', {
                  label: strings.AdminGroupID
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
