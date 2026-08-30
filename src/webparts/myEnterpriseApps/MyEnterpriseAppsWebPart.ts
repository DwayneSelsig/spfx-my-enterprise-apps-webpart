import * as React from 'react';
import * as ReactDom from 'react-dom';
import { Version } from '@microsoft/sp-core-library';
import {
  type IPropertyPaneConfiguration,
  PropertyPaneTextField,
  PropertyPaneCheckbox,
  PropertyPaneToggle,
  PropertyPaneDropdown,
  PropertyPaneSlider
} from '@microsoft/sp-property-pane';
import { BaseClientSideWebPart } from '@microsoft/sp-webpart-base';
import { MSGraphClientV3 } from '@microsoft/sp-http';

import * as strings from 'MyEnterpriseAppsWebPartStrings';
import MyEnterpriseApps from './components/MyEnterpriseApps';
import { IMyEnterpriseAppsProps } from './components/IMyEnterpriseAppsProps';
import { defaultApps } from './assets/DefaultApps';

type LayoutPreset = 'small' | 'normal' | 'large' | 'huge';
type LayoutPresetSelection = LayoutPreset | 'custom';

/**
 * WebPart properties interface
 */
export interface IMyEnterpriseAppsWebPartProps {
  title: string;
  sortOrder: string;
  showHiddenApps: boolean;
  showDefaultApps: boolean;
  defaultAppVisibility?: Record<string, boolean>;
  displayInternalNotes?: boolean;
  displayAppIdentifiers?: boolean;
  displayOAuthScopes?: boolean;
  enableDetailView?: boolean;
  iconSize?: number | 'small' | 'normal' | 'large' | 'huge';
  textSize?: number;
  appSpacing?: number;
  layoutPreset?: LayoutPresetSelection;
}

export default class MyEnterpriseAppsWebPart extends BaseClientSideWebPart<IMyEnterpriseAppsWebPartProps> {
  private graphClient!: MSGraphClientV3;

  private static readonly defaultIconSize: number = 48;
  private static readonly defaultTextSize: number = 11;
  private static readonly defaultAppSpacing: number = 12;

  private getLayoutForPreset(preset: LayoutPreset): { iconSize: number; textSize: number; appSpacing: number } {
    switch (preset) {
      case 'small':
        return { iconSize: 32, textSize: 9, appSpacing: 8 };
      case 'large':
        return { iconSize: 60, textSize: 12, appSpacing: 14 };
      case 'huge':
        return { iconSize: 80, textSize: 14, appSpacing: 16 };
      case 'normal':
      default:
        return {
          iconSize: MyEnterpriseAppsWebPart.defaultIconSize,
          textSize: MyEnterpriseAppsWebPart.defaultTextSize,
          appSpacing: MyEnterpriseAppsWebPart.defaultAppSpacing
        };
    }
  }

  private getLegacyLayout(): { iconSize: number; textSize: number; appSpacing: number } {
    switch (this.properties.iconSize) {
      case 'small':
        return this.getLayoutForPreset('small');
      case 'large':
        return this.getLayoutForPreset('large');
      case 'huge':
        return this.getLayoutForPreset('huge');
      case 'normal':
      default:
        return this.getLayoutForPreset('normal');
    }
  }

  private migrateLegacyLayout(): void {
    const legacyPreset = typeof this.properties.iconSize === 'string'
      ? this.properties.iconSize
      : undefined;
    const legacyLayout = this.getLegacyLayout();

    if (typeof this.properties.iconSize !== 'number') {
      this.properties.iconSize = legacyLayout.iconSize;
    }
    if (typeof this.properties.textSize !== 'number') {
      this.properties.textSize = legacyLayout.textSize;
    }
    if (typeof this.properties.appSpacing !== 'number') {
      this.properties.appSpacing = legacyLayout.appSpacing;
    }
    if (legacyPreset) {
      this.properties.layoutPreset = legacyPreset;
    }
    if (typeof this.properties.displayAppIdentifiers !== 'boolean') {
      this.properties.displayAppIdentifiers = true;
    }
    if (typeof this.properties.displayOAuthScopes !== 'boolean') {
      this.properties.displayOAuthScopes = true;
    }
    if (typeof this.properties.enableDetailView !== 'boolean') {
      this.properties.enableDetailView = true;
    }
  }

  private getSelectedLayoutPreset(): LayoutPresetSelection {
    return this.properties.layoutPreset || 'custom';
  }

  private getDefaultAppVisibilityKey(appName: string): string {
    return appName.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '');
  }

  private isDefaultAppVisible(appName: string): boolean {
    const key = this.getDefaultAppVisibilityKey(appName);
    return this.properties.defaultAppVisibility?.[key] !== false;
  }

  private getVisibleDefaultAppNames(): string[] {
    return defaultApps
      .filter(defaultApp => this.isDefaultAppVisible(defaultApp.name))
      .map(defaultApp => defaultApp.name);
  }

  protected onPropertyPaneFieldChanged(propertyPath: string, oldValue: unknown, newValue: unknown): void {
    if (propertyPath === 'layoutPreset' && newValue !== 'custom') {
      const preset = newValue as LayoutPreset;
      const layout = this.getLayoutForPreset(preset);
      this.properties.layoutPreset = preset;
      this.properties.iconSize = layout.iconSize;
      this.properties.textSize = layout.textSize;
      this.properties.appSpacing = layout.appSpacing;
    } else if (
      (propertyPath === 'iconSize' || propertyPath === 'textSize' || propertyPath === 'appSpacing') &&
      oldValue !== newValue
    ) {
      this.properties.layoutPreset = 'custom';
    }

    super.onPropertyPaneFieldChanged(propertyPath, oldValue, newValue);
    this.context.propertyPane.refresh();
  }

  public render(): void {
    const element: React.ReactElement<IMyEnterpriseAppsProps> = React.createElement(
      MyEnterpriseApps,
      {
        title: this.properties.title,
        sortOrder: this.properties.sortOrder,
        showHiddenApps: this.properties.showHiddenApps,
        showDefaultApps: this.properties.showDefaultApps,
        visibleDefaultAppNames: this.getVisibleDefaultAppNames(),
        displayInternalNotes: !!this.properties.displayInternalNotes,
        displayAppIdentifiers: this.properties.displayAppIdentifiers !== false,
        displayOAuthScopes: this.properties.displayOAuthScopes !== false,
        enableDetailView: this.properties.enableDetailView !== false,
        iconSize: this.properties.iconSize as number,
        textSize: this.properties.textSize as number,
        appSpacing: this.properties.appSpacing as number,
        hasTeamsContext: !!this.context.sdks.microsoftTeams,
        graphClient: this.graphClient
      }
    );

    ReactDom.render(element, this.domElement);
  }

  protected async onInit(): Promise<void> {
    await super.onInit();
    this.migrateLegacyLayout();
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
                PropertyPaneDropdown('layoutPreset', {
                  label: strings.LayoutPresetFieldLabel,
                  selectedKey: this.getSelectedLayoutPreset(),
                  options: [
                    { key: 'small', text: strings.LayoutPresetSmall },
                    { key: 'normal', text: strings.LayoutPresetNormal },
                    { key: 'large', text: strings.LayoutPresetLarge },
                    { key: 'huge', text: strings.LayoutPresetHuge },
                    { key: 'custom', text: strings.LayoutPresetCustom }
                  ]
                }),
                PropertyPaneSlider('iconSize', {
                  label: `${strings.IconSizeFieldLabel} (${this.properties.iconSize}px)`,
                  min: 12,
                  max: 256,
                  step: 4,
                  value: this.properties.iconSize as number,
                  showValue: false
                }),
                PropertyPaneSlider('textSize', {
                  label: `${strings.TextSizeFieldLabel} (${this.properties.textSize}px)`,
                  min: 6,
                  max: 48,
                  step: 1,
                  value: this.properties.textSize,
                  showValue: false
                }),
                PropertyPaneSlider('appSpacing', {
                  label: `${strings.AppSpacingFieldLabel} (${this.properties.appSpacing}px)`,
                  min: 0,
                  max: 64,
                  step: 2,
                  value: this.properties.appSpacing,
                  showValue: false
                })
              ]
            },
            {
              isGroupNameHidden: true,
              groupFields: [
                PropertyPaneToggle('enableDetailView', {
                  label: strings.DetailViewGroupName,
                  inlineLabel: true,
                  checked: this.properties.enableDetailView !== false,
                  ariaLabel: strings.EnableDetailViewLabel
                }),
                PropertyPaneCheckbox('displayInternalNotes', {
                  text: strings.DisplayInternalNotesLabel,
                  disabled: this.properties.enableDetailView === false
                }),
                PropertyPaneCheckbox('displayAppIdentifiers', {
                  text: strings.DisplayAppIdentifiersLabel,
                  disabled: this.properties.enableDetailView === false
                }),
                PropertyPaneCheckbox('displayOAuthScopes', {
                  text: strings.DisplayOAuthScopesLabel,
                  disabled: this.properties.enableDetailView === false
                })
              ]
            },
            {
              isGroupNameHidden: true,
              groupFields: [
                PropertyPaneCheckbox('showHiddenApps', {
                  text: strings.ShowHiddenAppsLabel
                })
              ]
            },
            {
              isGroupNameHidden: true,
              groupFields: [
                PropertyPaneToggle('showDefaultApps', {
                  label: strings.DefaultAppsGroupName,
                  inlineLabel: true,
                  checked: this.properties.showDefaultApps !== false,
                  ariaLabel: strings.ShowDefaultAppsLabel
                }),
                ...defaultApps.map(defaultApp => PropertyPaneCheckbox(
                  `defaultAppVisibility.${this.getDefaultAppVisibilityKey(defaultApp.name)}`,
                  {
                    text: defaultApp.name,
                    checked: this.isDefaultAppVisible(defaultApp.name),
                    disabled: this.properties.showDefaultApps === false
                  }
                ))
              ]
            }
          ]
        }
      ]
    };
  }
}
