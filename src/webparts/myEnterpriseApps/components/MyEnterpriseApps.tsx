import * as React from 'react';
import { IconButton, SearchBox, type ISearchBox } from '@fluentui/react';
import styles from './MyEnterpriseApps.module.scss';
import type {
  IMyEnterpriseAppsProps,
  IMyEnterpriseAppsState,
  IAppData,
  IAppRoleAssignment,
  IServicePrincipalInfo
} from './IMyEnterpriseAppsProps';
import { defaultApps } from '../assets/DefaultApps';
import * as strings from 'MyEnterpriseAppsWebPartStrings';

export default class MyEnterpriseApps extends React.Component<IMyEnterpriseAppsProps, IMyEnterpriseAppsState> {
  private readonly searchBoxRef = React.createRef<ISearchBox>();
  private detailTransitionTimer: number | undefined;
  private detailRequestId = 0;
  
  constructor(props: IMyEnterpriseAppsProps) {
    super(props);
    this.state = {
      apps: [],
      isLoading: true,
      error: undefined,
      filterQuery: '',
      isFilterOpen: false,
      selectedApp: undefined,
      isDetailTransitioning: false,
      isReturningToResults: false,
      isDetailDismissed: false
    };
  }

  public componentWillUnmount(): void {
    this.cancelDetailTransition();
    this.detailRequestId++;
  }

  public componentDidMount(): void {
    this.loadApps().catch(error => {
      console.error('Error loading apps:', error);
      this.setState({ error: error.message, isLoading: false });
    });
  }

  public componentDidUpdate(prevProps: IMyEnterpriseAppsProps): void {
    if (prevProps.enableDetailView && !this.props.enableDetailView && this.state.selectedApp) {
      this.cancelDetailTransition();
      this.setState({
        selectedApp: undefined,
        isDetailTransitioning: false,
        isReturningToResults: false,
        isDetailDismissed: false
      });
    }
    if (!prevProps.enableDetailView && this.props.enableDetailView) {
      this.openDetailForExactMatch();
    }

    // Reload if sorting, app visibility, or default app selection changes
    if (prevProps.sortOrder !== this.props.sortOrder ||
        prevProps.showHiddenApps !== this.props.showHiddenApps ||
        prevProps.showDefaultApps !== this.props.showDefaultApps ||
        prevProps.visibleDefaultAppNames.join('|') !== this.props.visibleDefaultAppNames.join('|')) {
      this.cancelDetailTransition();
      this.detailRequestId++;
      this.setState({
        selectedApp: undefined,
        isDetailTransitioning: false,
        isReturningToResults: false,
        isDetailDismissed: false
      });
      this.loadApps().catch(error => {
        console.error('Error loading apps:', error);
        this.setState({ error: error.message, isLoading: false });
      });
    }
  }

  /**
   * Generate a default SVG icon with the first letter of the app name
   */
  private generateDefaultIcon(appName: string): string {
    const firstLetter = appName.charAt(0).toUpperCase();
    return `data:image/svg+xml,%3Csvg xmlns='http://www.w3.org/2000/svg' viewBox='0 0 48 48'%3E%3Crect width='48' height='48' rx='8' fill='%230078d4'/%3E%3Ctext x='24' y='32' text-anchor='middle' fill='white' font-size='22' font-family='Segoe UI, sans-serif' font-weight='600'%3E${firstLetter}%3C/text%3E%3C/svg%3E`;
  }

  private normalizeName(name: string): string {
    return (name || '').trim().toLowerCase();
  }

  private openFilter = (): void => {
    this.setState({ isFilterOpen: true }, () => {
      this.searchBoxRef.current?.focus();
    });
  };

  private closeFilter = (): void => {
    this.cancelDetailTransition();
    this.detailRequestId++;
    this.setState({
      filterQuery: '',
      isFilterOpen: false,
      selectedApp: undefined,
      isDetailTransitioning: false,
      isReturningToResults: false,
      isDetailDismissed: false
    });
  };

  private onFilterEscape = (event?: { preventDefault: () => void; stopPropagation: () => void }): void => {
    event?.preventDefault();
    event?.stopPropagation();
    this.closeFilter();
  };

  private onFilterChange = (_event?: React.ChangeEvent<HTMLInputElement>, newValue?: string): void => {
    const previousResultCount = this.getFilteredApps().length;
    const nextFilterQuery = newValue || '';
    const nextMatchingApps = this.getFilteredApps(nextFilterQuery);
    const nextResultCount = nextMatchingApps.length;
    const delayBeforeDetail = previousResultCount !== 1 && nextResultCount === 1 ? 250 : 0;
    const selectedApp = this.state.selectedApp;
    const wasShowingDetail = !!selectedApp;
    const selectedAppKey = selectedApp?.resourceId || this.normalizeName(selectedApp?.name || '');
    const nextAppKey = nextMatchingApps[0]?.resourceId || this.normalizeName(nextMatchingApps[0]?.name || '');
    const remainsOnSelectedApp = wasShowingDetail && nextResultCount === 1 && selectedAppKey === nextAppKey;

    this.cancelDetailTransition();
    this.detailRequestId++;

    if (remainsOnSelectedApp) {
      this.setState({
        filterQuery: nextFilterQuery,
        isDetailTransitioning: false,
        isReturningToResults: false,
        isDetailDismissed: false
      });
      return;
    }

    const nextState = {
      filterQuery: nextFilterQuery,
      selectedApp: undefined,
      isDetailTransitioning: false,
      isReturningToResults: false,
      isDetailDismissed: false
    };
    const afterFilterChange = (): void => this.openDetailForExactMatch(delayBeforeDetail);

    if (wasShowingDetail && this.startViewTransition(complete => {
      this.setState(nextState, () => {
        complete();
        window.requestAnimationFrame(afterFilterChange);
      });
    })) {
      return;
    }

    this.setState(nextState, afterFilterChange);
  };

  private cancelDetailTransition(): void {
    if (this.detailTransitionTimer !== undefined) {
      window.clearTimeout(this.detailTransitionTimer);
      this.detailTransitionTimer = undefined;
    }
  }

  private prefersReducedMotion(): boolean {
    return typeof window !== 'undefined' &&
      typeof window.matchMedia === 'function' &&
      window.matchMedia('(prefers-reduced-motion: reduce)').matches;
  }

  private startViewTransition(update: (complete: () => void) => void): boolean {
    if (this.prefersReducedMotion() || typeof document === 'undefined') {
      return false;
    }

    const transitionDocument = document as Partial<Pick<Document, 'startViewTransition'>>;
    if (typeof transitionDocument.startViewTransition !== 'function') {
      return false;
    }

    transitionDocument.startViewTransition(() => new Promise<void>(resolve => update(resolve)));
    return true;
  }

  private getFilteredApps(query: string = this.state.filterQuery): IAppData[] {
    const normalizedQuery = this.normalizeName(query);
    return normalizedQuery
      ? this.state.apps.filter(app => this.normalizeName(app.name).indexOf(normalizedQuery) !== -1)
      : this.state.apps;
  }

  private openDetailForExactMatch = (delayBeforeDetail: number = 0): void => {
    const { isLoading, isDetailDismissed } = this.state;
    const hasSearchQuery = this.normalizeName(this.state.filterQuery).length > 0;
    const matchingApps = this.getFilteredApps();

    if (!this.props.enableDetailView || isLoading || isDetailDismissed || !hasSearchQuery || matchingApps.length !== 1) {
      return;
    }

    const selectedApp = matchingApps[0];
    const showDetail = (): void => {
      this.detailTransitionTimer = undefined;
      this.setState({
        selectedApp,
        isDetailTransitioning: false,
        isDetailDismissed: false
      });
    };

    const beginDetailTransition = (): void => {
      this.detailTransitionTimer = undefined;

      if (this.startViewTransition(complete => {
        this.setState({
          selectedApp,
          isDetailTransitioning: false,
          isReturningToResults: false,
          isDetailDismissed: false
        }, complete);
      })) {
        return;
      }

      if (this.prefersReducedMotion()) {
        showDetail();
        return;
      }

      this.setState({ isDetailTransitioning: true });
      this.detailTransitionTimer = window.setTimeout(showDetail, 160);
    };

    if (delayBeforeDetail > 0 && !this.prefersReducedMotion()) {
      this.detailTransitionTimer = window.setTimeout(beginDetailTransition, delayBeforeDetail);
      return;
    }

    beginDetailTransition();
  };

  private returnToResults = (): void => {
    this.cancelDetailTransition();
    this.detailRequestId++;
    const showResults = (): void => {
      this.detailTransitionTimer = undefined;
      this.setState({
        selectedApp: undefined,
        isDetailTransitioning: false,
        isReturningToResults: false,
        isDetailDismissed: true
      });
    };

    if (this.startViewTransition(complete => {
      this.setState({
        selectedApp: undefined,
        isDetailTransitioning: false,
        isReturningToResults: false,
        isDetailDismissed: true
      }, complete);
    })) {
      return;
    }

    if (this.prefersReducedMotion()) {
      showResults();
      return;
    }

    this.setState({ isReturningToResults: true });
    this.detailTransitionTimer = window.setTimeout(showResults, 180);
  };

  /**
   * Load apps from Microsoft Graph API
   */
  private async loadApps(): Promise<void> {
    try {
      const { graphClient } = this.props;

      // Get app role assignments from Graph API
      const response = await graphClient
        .api('/me/appRoleAssignments')
        .get();

      const assignments: IAppRoleAssignment[] = response.value || [];
      const includeDefaults = this.props.showDefaultApps ?? true;
      const visibleDefaultAppKeys = includeDefaults
        ? this.props.visibleDefaultAppNames.map(appName => this.normalizeName(appName))
        : [];
      const defaultAppKeys = defaultApps.map(defaultApp => this.normalizeName(defaultApp.name));

      // Build a Map with all apps (keyed by normalized name)
      const allAppsMap = new Map<string, IAppData>();
      
      assignments.forEach(assignment => {
        const defaultIcon = this.generateDefaultIcon(assignment.resourceDisplayName);
        const key = this.normalizeName(assignment.resourceDisplayName);

        if (defaultAppKeys.indexOf(key) !== -1 && visibleDefaultAppKeys.indexOf(key) === -1) {
          return;
        }

        allAppsMap.set(key, {
          name: assignment.resourceDisplayName,
          url: '',
          iconUrl: defaultIcon,
          resourceId: assignment.resourceId,
          isHidden: false,
          isLoaded: false,
          isDefaultApp: false
        });
      });

      if (includeDefaults) {
        defaultApps.forEach(defaultApp => {
          if (this.props.visibleDefaultAppNames.indexOf(defaultApp.name) === -1) {
            return;
          }

          const key = this.normalizeName(defaultApp.name);
          const existing = allAppsMap.get(key);
          const merged: IAppData = {
            name: defaultApp.name,
            url: defaultApp.url,
            iconUrl: defaultApp.icon,
            resourceId: existing?.resourceId || '',
            isHidden: existing?.isHidden ?? false,
            isLoaded: existing ? existing.isLoaded : true,
            isDefaultApp: true
          };

          if (existing) {
            allAppsMap.set(key, { ...existing, ...merged });
          } else {
            allAppsMap.set(key, merged);
          }
        });
      }

      if (allAppsMap.size === 0) {
        this.setState({ apps: [], isLoading: false });
        return;
      }

      // Sort apps according to custom logic
      const appsArray = this.sortApps(allAppsMap);

      // Set initial state with sorted apps
      this.setState({ apps: appsArray });

      // Load service principal details asynchronously
      await this.loadServicePrincipalDetails(appsArray);

    } catch (error) {
      console.error('Error loading apps:', error);
      throw error;
    }
    finally {
      this.setState({ isLoading: false }, this.openDetailForExactMatch);
    }
  }

  /**
   * Sort apps according to custom sort order and alphabetically
   */
  private sortApps(appsMap: Map<string, IAppData>): IAppData[] {
    const appsArray: IAppData[] = [];
    appsMap.forEach((value) => {
      appsArray.push(value);
    });

    // Parse custom sort order
    const customOrder: string[] = [];
    if (this.props.sortOrder) {
      const lines = this.props.sortOrder.split('\n');
      lines.forEach(line => {
        const trimmed = line.trim();
        if (trimmed) {
          customOrder.push(trimmed.toLowerCase());
        }
      });
    }

    // Custom sorting: first by custom order, then alphabetically
    appsArray.sort((a, b) => {
      const aName = a.name.toLowerCase();
      const bName = b.name.toLowerCase();

      // Check if names match custom order items
      let aIndex = -1;
      let bIndex = -1;

      for (let i = 0; i < customOrder.length; i++) {
        if (aName.indexOf(customOrder[i]) !== -1 && aIndex === -1) {
          aIndex = i;
        }
        if (bName.indexOf(customOrder[i]) !== -1 && bIndex === -1) {
          bIndex = i;
        }
      }

      // Both in custom order: sort by custom order
      if (aIndex !== -1 && bIndex !== -1) {
        return aIndex - bIndex;
      }

      // Only a in custom order: a comes first
      if (aIndex !== -1) {
        return -1;
      }

      // Only b in custom order: b comes first
      if (bIndex !== -1) {
        return 1;
      }

      // Neither in custom order: sort alphabetically
      return aName.localeCompare(bName);
    });

    return appsArray;
  }

  /**
   * Load service principal details and update app state
   */
  private async loadServicePrincipalDetails(apps: IAppData[]): Promise<void> {
    const { graphClient, showHiddenApps } = this.props;

    const loadPromises = apps.map(async (app) => {
      try {
        if (!app.resourceId) {
          return { ...app, isLoaded: true } as IAppData;
        }

        const spInfo: IServicePrincipalInfo = await graphClient
          .api(`/servicePrincipals/${app.resourceId}`)
          .select('id,appId,appOwnerOrganizationId,displayName,appDescription,notes,homepage,publisherName,verifiedPublisher,preferredSingleSignOnMode,info,tags,oauth2PermissionScopes')
          .get();

        // Construct login URL
        const loginUrl = spInfo.appOwnerOrganizationId
          ? `https://launcher.myapps.microsoft.com/api/signin/${spInfo.appId}?tenantId=${spInfo.appOwnerOrganizationId}`
          : app.url;

        // Check if app has HideApp tag
        const hasHideAppTag = spInfo.tags && Array.isArray(spInfo.tags) && spInfo.tags.indexOf('HideApp') !== -1;

        // Update app data
        return {
          ...app,
          url: app.isDefaultApp ? app.url : loginUrl,
          iconUrl: spInfo.info?.logoUrl || app.iconUrl,
          isHidden: hasHideAppTag === true,
          isLoaded: true,
          servicePrincipal: spInfo
        } as IAppData;
      } catch (error) {
        console.warn(`Could not load details for ${app.name}:`, error);
        return { ...app, isLoaded: true };
      }
    });

    const results = await Promise.all(loadPromises);

    // Deduplicate by normalized name to avoid double-defaults during rapid refreshes
    const uniqueByName = new Map<string, IAppData>();
    results.forEach(app => {
      uniqueByName.set(this.normalizeName(app.name), app);
    });
    const dedupedApps = Array.from(uniqueByName.values());
    
    // Filter apps based on showHiddenApps setting
    const filteredApps = showHiddenApps 
      ? dedupedApps 
      : dedupedApps.filter(app => !app.isHidden);

    this.setState({ apps: filteredApps }, this.openDetailForExactMatch);
  }

  /**
   * Render skeleton loading items
   */
  private renderSkeletonItems(): React.ReactElement[] {
    const items: React.ReactElement[] = [];
    for (let i = 0; i < 11; i++) {
      items.push(
        <div key={`skeleton-${i}`} className={`${styles.appItem} ${styles.skeleton}`}>
          <div className={`${styles.appIcon} ${styles.skeletonIcon}`} />
          <div className={`${styles.appName} ${styles.skeletonText}`} />
        </div>
      );
    }
    return items;
  }

  /**
   * Render a single app item
   */
  private renderAppItem(app: IAppData, isTransitionCandidate: boolean = false): React.ReactElement {
    const appClasses = `${styles.appItem}${app.isHidden ? ` ${styles.hiddenApp}` : ''}${isTransitionCandidate ? ` ${styles.transitionCandidate}` : ''}`;

    const content = (
      <>
        <div className={styles.appIcon}>
          <img src={app.iconUrl} alt={app.name} />
        </div>
        <div className={styles.appName}>{app.name}</div>
      </>
    );

    return (
      <div 
        key={app.resourceId || app.name} 
        className={appClasses}
        data-resource-id={app.resourceId}
        data-app-name={app.name}
        data-hidden={app.isHidden ? 'true' : undefined}
      >
        {app.url ? (
          <a href={app.url} target="_blank" rel="noopener noreferrer">
            {content}
          </a>
        ) : (
          content
        )}
      </div>
    );
  }

  private getSsoLabel(mode: string | undefined): string | undefined {
    if (!mode) {
      return undefined;
    }

    const labels: { [key: string]: string } = {
      saml: 'SAML',
      password: strings.PasswordSso,
      oidc: 'OpenID Connect',
      linked: strings.LinkedSso,
      notSupported: strings.SsoNotSupported
    };
    return labels[mode] || mode;
  }

  private renderOAuthScopes(servicePrincipal: IServicePrincipalInfo | undefined): React.ReactElement | null {
    const scopes = (servicePrincipal?.oauth2PermissionScopes || []).filter(scope => scope.isEnabled !== false);

    if (scopes.length === 0) {
      return null;
    }

    return (
      <section className={styles.detailInfoBlock}>
        <h3>{strings.OAuthScopes}</h3>
        <ul className={styles.scopesList}>
          {scopes.map(scope => (
            <li key={scope.id || scope.value}>
              <span className={styles.scopeName}>{scope.value}</span>
              {(scope.userConsentDescription || scope.adminConsentDescription) && (
                <span className={styles.scopeDescription}>
                  {scope.userConsentDescription || scope.adminConsentDescription}
                </span>
              )}
            </li>
          ))}
        </ul>
      </section>
    );
  }

  private renderAppDetail(app: IAppData): React.ReactElement {
    const servicePrincipal = app.servicePrincipal;
    const publisher = servicePrincipal?.publisherName;
    const verifiedPublisher = servicePrincipal?.verifiedPublisher?.displayName;
    const ssoMode = this.getSsoLabel(servicePrincipal?.preferredSingleSignOnMode);
    const termsUrl = servicePrincipal?.info?.termsOfServiceUrl;
    const homepage = servicePrincipal?.homepage;
    const notes = servicePrincipal?.notes?.trim();
    const hasResources = !!homepage || !!termsUrl || !!ssoMode;
    const hasOAuthScopes = servicePrincipal?.oauth2PermissionScopes?.some(scope => scope.isEnabled !== false) === true;
    const showAppIdentifiers = this.props.displayAppIdentifiers;
    const showOAuthScopes = this.props.displayOAuthScopes && hasOAuthScopes;

    return (
      <div className={styles.detailView} aria-live="polite">
        <div className={styles.detailToolbar}>
          <IconButton
            iconProps={{ iconName: 'Back' }}
            ariaLabel={strings.BackToResults}
            title={strings.BackToResults}
            onClick={this.returnToResults}
          />
          <button type="button" className={styles.backButton} onClick={this.returnToResults}>
            {strings.BackToResults}
          </button>
        </div>

        <article className={styles.detailCard}>
          <div className={styles.detailLayout}>
            <aside className={styles.detailAside}>
              <div className={styles.detailIcon}>
                <img src={app.iconUrl} alt={app.name} />
              </div>
              {app.url && (
                <a className={styles.openAppButton} href={app.url} target="_blank" rel="noopener noreferrer">
                  {strings.OpenApp}
                </a>
              )}
            </aside>

            <div className={styles.detailContent}>
              <header className={styles.detailHeader}>
                <div className={styles.detailIdentity}>
                  <h2 className={styles.detailTitle}>{servicePrincipal?.displayName || app.name}</h2>
                  {publisher && <p className={styles.publisher}>{publisher}</p>}
                  {verifiedPublisher && <p className={styles.verifiedPublisher}>{strings.VerifiedPublisher}: {verifiedPublisher}</p>}
                </div>
                {this.props.displayInternalNotes && notes && <p className={styles.appNotes}>{notes}</p>}
                {hasResources && (
                  <section className={styles.detailInfoBlock}>
                    <h3>{strings.Resources}</h3>
                    <ul className={styles.resourceList}>
                      {homepage && (
                        <li><a href={homepage} target="_blank" rel="noopener noreferrer">{strings.Homepage}</a></li>
                      )}
                      {termsUrl && (
                        <li><a href={termsUrl} target="_blank" rel="noopener noreferrer">{strings.OpenTermsOfService}</a></li>
                      )}
                      {ssoMode && (
                        <li><span>{strings.SingleSignOn}</span><strong>{ssoMode}</strong></li>
                      )}
                    </ul>
                  </section>
                )}
              </header>

              {servicePrincipal?.appDescription && (
                <p className={styles.appDescription}>{servicePrincipal.appDescription}</p>
              )}

              {(showAppIdentifiers || showOAuthScopes) && (
                <div className={styles.detailInfoGrid}>
                  {showAppIdentifiers && (
                    <div className={styles.metadata}>
                      <div><span>{strings.AppId}</span><code>{servicePrincipal?.appId || strings.NotAvailable}</code></div>
                      <div><span>{strings.ServicePrincipalId}</span><code>{app.resourceId || strings.NotAvailable}</code></div>
                    </div>
                  )}
                  {showOAuthScopes && this.renderOAuthScopes(servicePrincipal)}
                </div>
              )}
            </div>
          </div>
        </article>
      </div>
    );
  }

  public render(): React.ReactElement<IMyEnterpriseAppsProps> {
    const { title, hasTeamsContext, iconSize, textSize, appSpacing } = this.props;
    const { apps, isLoading, error, filterQuery, isFilterOpen, selectedApp, isDetailTransitioning } = this.state;
    const displayTitle = title || strings.DefaultTitle;
    const normalizedFilterQuery = this.normalizeName(filterQuery);
    const filteredApps = normalizedFilterQuery
      ? apps.filter(app => this.normalizeName(app.name).indexOf(normalizedFilterQuery) !== -1)
      : apps;
    const hasNoSearchResults = normalizedFilterQuery.length > 0 && filteredApps.length === 0;
    const hasExactSearchResult = this.props.enableDetailView && normalizedFilterQuery.length > 0 && filteredApps.length === 1;
    const layoutStyle = {
      '--app-icon-size': `${iconSize}px`,
      '--app-text-size': `${textSize}px`,
      '--app-spacing': `${appSpacing}px`,
      '--app-tile-min-width': `${Math.max(90, iconSize + 24)}px`
    } as React.CSSProperties;

    return (
      <section className={`${styles.myEnterpriseApps} ${hasTeamsContext ? styles.teams : ''}`} style={layoutStyle}>
        <div className={styles.container}>
          <div className={styles.header}>
            <h2>
                {displayTitle}
            </h2>
            <a 
              href="https://myapps.microsoft.com/" 
              target="_blank" 
              rel="noopener noreferrer" 
              className={styles.allAppsLink}
            >
              {strings.AllAppsLabel}
            </a>
          </div>

          <div className={styles.filterBar}>
            {!isFilterOpen ? (
              <IconButton
                iconProps={{ iconName: 'Filter' }}
                ariaLabel={strings.FilterAppsLabel}
                title={strings.FilterAppsLabel}
                onClick={this.openFilter}
              />
            ) : (
              <>
                <SearchBox
                  componentRef={this.searchBoxRef}
                  className={styles.filterSearchBox}
                  value={filterQuery}
                  placeholder={strings.FilterAppsPlaceholder}
                  ariaLabel={strings.FilterAppsLabel}
                  clearButtonProps={{ ariaLabel: strings.ClearFilterLabel }}
                  onChange={this.onFilterChange}
                  onEscape={this.onFilterEscape}
                />
                <IconButton
                  iconProps={{ iconName: 'Cancel' }}
                  ariaLabel={strings.CloseFilterLabel}
                  title={strings.CloseFilterLabel}
                  onClick={this.closeFilter}
                />
              </>
            )}
          </div>
          
          {selectedApp ? (
            <div className={this.state.isReturningToResults ? styles.detailViewLeaving : undefined}>
              {this.renderAppDetail(selectedApp)}
            </div>
          ) : (
            <div className={`${styles.appsList}${isDetailTransitioning ? ` ${styles.appsListFading}` : ''}`}>
              {error && (
                <div className={styles.errorMessage}>
                  {strings.ErrorLoading}: {error}
                </div>
              )}

              {isLoading && this.renderSkeletonItems()}

              {!isLoading && !error && hasNoSearchResults && (
                <div className={styles.noAppsMessage}>
                  {strings.NoFilterResults}
                </div>
              )}

              {!isLoading && !error && !hasNoSearchResults && apps.length === 0 && (
                <div className={styles.noAppsMessage}>
                  {strings.NoAppsFound}
                </div>
              )}

              {!isLoading && !error && !hasNoSearchResults && filteredApps.map(app => this.renderAppItem(app, hasExactSearchResult))}
            </div>
          )}
        </div>
      </section>
    );
  }
}
