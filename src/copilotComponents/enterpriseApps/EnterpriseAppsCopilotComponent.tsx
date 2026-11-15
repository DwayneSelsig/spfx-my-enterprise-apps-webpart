import * as React from 'react';
import { createRoot, type Root } from 'react-dom/client';
import { ThemeProvider, createTheme } from '@fluentui/react';
import { BaseCopilotComponent } from '@microsoft/sp-copilot-component';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import MyEnterpriseApps from '../../webparts/myEnterpriseApps/components/MyEnterpriseApps';
import { defaultApps } from '../../webparts/myEnterpriseApps/assets/DefaultApps';
import type { IMyEnterpriseAppsProps } from '../../webparts/myEnterpriseApps/components/IMyEnterpriseAppsProps';
import type { IEnterpriseAppsCopilotComponentProperties } from './EnterpriseAppsCopilotComponentProperties';
import * as strings from 'MyEnterpriseAppsWebPartStrings';

/** Adapts the existing app picker to Copilot; data access stays in MyEnterpriseApps. */
export default class EnterpriseAppsCopilotComponent extends BaseCopilotComponent<IEnterpriseAppsCopilotComponentProperties> {
  private root: Root | undefined;
  private graphClientPromise: Promise<MSGraphClientV3> | undefined;
  private linkError = false;
  private disposed = false;

  private getGraphClient = (): Promise<MSGraphClientV3> => {
    if (!this.graphClientPromise) {
      this.graphClientPromise = this.context.msGraphClientFactory.getClient('3').catch(error => {
        this.graphClientPromise = undefined;
        throw error;
      });
    }
    return this.graphClientPromise;
  };

  // Never use an iframe/session identifier as the user's persistent cache identity.
  private getIdentity(value: unknown): string | undefined {
    const id = typeof value === 'string' ? value.trim() : undefined;
    return id && /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i.test(id)
      ? id : undefined;
  }

  private openLink = (event: React.MouseEvent<HTMLDivElement>): void => {
    const target = event.target as Element;
    const anchor = target.closest?.('a[href]') as HTMLAnchorElement | null;
    if (!anchor || event.defaultPrevented || event.button !== 0) {
      return;
    }
    // Let the Copilot host open links outside its sandbox, only after a user click.
    event.preventDefault();
    this.linkError = false;
    this.context.copilotBridge.openLinkAsync(anchor.href).then(result => {
      this.linkError = result.isError === true;
      if (!this.disposed) this.render();
    }).catch(() => {
      this.linkError = true;
      if (!this.disposed) this.render();
    });
  };

  protected render(): void {
    const dark = this.hostContext.theme === 'dark';
    const theme = createTheme(dark ? {
      isInverted: true,
      palette: { themePrimary: '#479ef5', white: '#1f1f1f', black: '#ffffff',
        neutralPrimary: '#ffffff', neutralSecondary: '#d6d6d6', neutralLighter: '#292929' }
    } : {});
    const aad = this.context.pageContext.aadInfo;
    const props: IMyEnterpriseAppsProps = {
      title: strings.DefaultTitle,
      sortOrder: '',
      showHiddenApps: false,
      showOnlyAssignedApps: true,
      showDefaultApps: true,
      visibleDefaultAppNames: defaultApps.map(app => app.name),
      enableCache: true,
      cacheDurationMinutes: 30,
      tenantId: this.getIdentity(aad?.tenantId?.toString()),
      userId: this.getIdentity(aad?.userId?.toString()),
      isPropertyPaneOpen: false,
      isEditMode: false,
      displayInternalNotes: false,
      displayAppIdentifiers: true,
      displayOAuthScopes: true,
      enableDetailView: true,
      iconSize: 48,
      textSize: 11,
      appSpacing: 12,
      themePrimary: theme.palette.themePrimary,
      bodyBackground: theme.semanticColors.bodyBackground || theme.palette.white,
      themePrimaryTextColor: dark ? '#000000' : '#ffffff',
      hasTeamsContext: false,
      getGraphClient: this.getGraphClient,
      initialSearchQuery: typeof this.properties.query === 'string' ? this.properties.query : ''
    };
    const style = {
      '--bodyText': theme.semanticColors.bodyText,
      '--bodySubtext': theme.semanticColors.bodySubtext,
      '--white': theme.palette.white,
      '--link': theme.semanticColors.link,
      '--linkHovered': theme.semanticColors.linkHovered,
      '--themePrimary': theme.palette.themePrimary,
      '--themeDark': theme.palette.themeDark
    } as React.CSSProperties;
    this.root = this.root ?? createRoot(this.context.domElement);
    this.root.render(
      <ThemeProvider theme={theme}>
        <div style={style} onClick={this.openLink}>
          {this.linkError && <div role="alert">{strings.CopilotOpenLinkError}</div>}
          <MyEnterpriseApps {...props} />
        </div>
      </ThemeProvider>
    );
  }

  protected onDispose(): void {
    this.disposed = true;
    this.root?.unmount();
    this.root = undefined;
  }
}
