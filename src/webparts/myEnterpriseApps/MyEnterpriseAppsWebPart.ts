import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneChoiceGroup
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'MyEnterpriseAppsWebPartStrings';
import MyEnterpriseApps from './components/MyEnterpriseApps';
import { IMyEnterpriseAppsProps } from './components/IMyEnterpriseAppsProps';

/**
 * WebPart properties interface
 */
export interface IMyEnterpriseAppsWebPartProps {
  title: string;
  sortOrder: string;
  showHiddenApps: boolean;
  iconSize: 'small' | 'normal' | 'large' | 'huge';
}

export default class MyEnterpriseAppsWebPart extends BaseClientSideWebPart<IMyEnterpriseAppsWebPartProps> {
  private graphClient!: MSGraphClientV3;

  public render(): void {
    const element: React.ReactElement<IMyEnterpriseAppsProps> = React.createElement(
      MyEnterpriseApps,
      {
        title: this.properties.title,
        sortOrder: this.properties.sortOrder,
        showHiddenApps: this.properties.showHiddenApps,
        iconSize: this.properties.iconSize || 'normal',
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        graphClient: this.graphClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    this.graphClient = await this.context.msGraphClientFactory.getClient('3');
    // Set a localized default title when empty
    if (!this.properties.title || this.properties.title.trim() === '') {
      this.properties.title = strings.DefaultTitle || this.properties.title;
    }
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
                PropertyPaneTextField('title', {
                  label: strings.TitleFieldLabel
                }),
                PropertyPaneTextField('sortOrder', {
                  label: strings.SortOrderFieldLabel,
                  description: strings.SortOrderFieldDescription,
                  multiline: true,
                  rows: 5
                }),
                PropertyPaneCheckbox('showHiddenApps', {
                  text: strings.ShowHiddenAppsLabel
                }),
                PropertyPaneChoiceGroup('iconSize', {
                  label: strings.IconSizeFieldLabel,
                  options: [
                    { key: 'small', text: strings.IconSizeSmall },
                    { key: 'normal', text: strings.IconSizeNormal },
                    { key: 'large', text: strings.IconSizeLarge },
                    { key: 'huge', text: strings.IconSizeHuge }
                  ]
                })
              ]
            }
          ]
        }
      ]
    };
  }
}
