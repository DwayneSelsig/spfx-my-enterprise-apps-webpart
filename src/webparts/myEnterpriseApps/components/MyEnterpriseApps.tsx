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
import {
  EnterpriseAppsCache,
  getEnterpriseAppsCacheKey,
  getEnterpriseAppsCacheSignature,
  type IEnterpriseAppsCacheConfiguration,
  normalizeCacheDuration
} from './EnterpriseAppsCache';

interface IGraphPage<T> {
  value?: T[];
  '@odata.nextLink'?: string;
}

interface IGraphRequest<T> {
  get(): Promise<IGraphPage<T>>;
}

interface IAppRoleAssignedTo {
  principalId: string;
  principalType: string;
}

interface IGraphBatchSubResponse {
  id: string;
  status: number;
  body?: IGraphPage<IAppRoleAssignedTo>;
}

interface IGraphBatchResponse {
  responses?: IGraphBatchSubResponse[];
}

interface IUnassignedAppsResult {
  apps: IServicePrincipalInfo[];
  isComplete: boolean;
}

interface ILoadedAppsResult {
  apps: IAppData[];
  isComplete: boolean;
}

export default class MyEnterpriseApps extends React.Component<IMyEnterpriseAppsProps, IMyEnterpriseAppsState> {
  private readonly searchBoxRef = React.createRef<ISearchBox>();
  private readonly cache = new EnterpriseAppsCache();
  private detailTransitionTimer: number | undefined;
  private detailRequestId = 0;
  private loadRequestId = 0;
  
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
    this.loadRequestId++;
  }

  public componentDidMount(): void {
    this.loadApps().catch(error => {
      this.handleLoadError(error);
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

    const sortOrderChanged = prevProps.sortOrder !== this.props.sortOrder;
    const cacheConfigurationChanged = getEnterpriseAppsCacheSignature(this.getCacheConfiguration(prevProps)) !==
      getEnterpriseAppsCacheSignature(this.getCacheConfiguration());
    const cacheIdentityChanged = prevProps.tenantId !== this.props.tenantId ||
      prevProps.userId !== this.props.userId;
    const cacheModeChanged = prevProps.enableCache !== this.props.enableCache;
    const managementContextChanged = prevProps.isPropertyPaneOpen !== this.props.isPropertyPaneOpen ||
      prevProps.isEditMode !== this.props.isEditMode;

    // Sorting is local and does not require another Graph request.
    if (sortOrderChanged && !cacheConfigurationChanged && !cacheIdentityChanged &&
        !cacheModeChanged && !managementContextChanged) {
      this.cancelDetailTransition();
      this.detailRequestId++;
      this.setState({
        apps: this.sortApps(this.state.apps),
        selectedApp: undefined,
        isDetailTransitioning: false,
        isReturningToResults: false,
        isDetailDismissed: false
      });
    } else if (cacheConfigurationChanged || cacheIdentityChanged || cacheModeChanged || managementContextChanged) {
      this.cancelDetailTransition();
      this.detailRequestId++;
      this.setState({
        apps: [],
        isLoading: true,
        error: undefined,
        selectedApp: undefined,
        isDetailTransitioning: false,
        isReturningToResults: false,
        isDetailDismissed: false
      });
      this.loadApps().catch(error => {
        this.handleLoadError(error);
      });
    }
  }

  private getCacheConfiguration(props: IMyEnterpriseAppsProps = this.props): IEnterpriseAppsCacheConfiguration {
    return {
      tenantId: props.tenantId,
      userId: props.userId,
      showHiddenApps: props.showHiddenApps === true,
      showDefaultApps: props.showDefaultApps !== false,
      visibleDefaultAppNames: Array.isArray(props.visibleDefaultAppNames) ? props.visibleDefaultAppNames : []
    };
  }

  private canUseCache(configuration: IEnterpriseAppsCacheConfiguration = this.getCacheConfiguration()): boolean {
    return this.props.enableCache !== false &&
      !this.props.isPropertyPaneOpen &&
      !this.props.isEditMode &&
      getEnterpriseAppsCacheKey(configuration.tenantId, configuration.userId) !== undefined;
  }

  private canWriteCache(
    configuration: IEnterpriseAppsCacheConfiguration,
    cacheWasAllowedWhenLoadStarted: boolean
  ): boolean {
    if (!cacheWasAllowedWhenLoadStarted || !this.canUseCache()) {
      return false;
    }

    const currentConfiguration = this.getCacheConfiguration();
    return getEnterpriseAppsCacheKey(currentConfiguration.tenantId, currentConfiguration.userId) ===
      getEnterpriseAppsCacheKey(configuration.tenantId, configuration.userId) &&
      getEnterpriseAppsCacheSignature(currentConfiguration) === getEnterpriseAppsCacheSignature(configuration);
  }

  private isCurrentLoad(loadRequestId: number): boolean {
    return loadRequestId === this.loadRequestId;
  }

  private getErrorMessage(error: unknown): string {
    if (error instanceof Error && error.message) {
      return error.message;
    }

    return typeof error === 'string' ? error : strings.ErrorLoading;
  }

  private handleLoadError(error: unknown): void {
    console.error('Error loading apps:', error);
    this.setState({ error: this.getErrorMessage(error), isLoading: false });
  }

  /**
   * Generate a default SVG icon with the first letter of the app name
   */
  private generateDefaultIcon(appName: string): string {
    const firstLetter = appName.charAt(0).toUpperCase();
    const backgroundColor = this.props.themePrimary;
    const textColor = this.props.themePrimaryTextColor;
    const svg = `
      <svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 48 48">
        <rect width="48" height="48" rx="8" fill="${backgroundColor}" />
        <text x="24" y="32" text-anchor="middle" fill="${textColor}" font-size="22" font-family="Segoe UI, sans-serif" font-weight="600">${firstLetter}</text>
      </svg>`;

    return `data:image/svg+xml,${encodeURIComponent(svg)}`;
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
    const loadRequestId = ++this.loadRequestId;
    const cacheConfiguration = this.getCacheConfiguration();
    const cacheWasAllowedWhenLoadStarted = this.canUseCache(cacheConfiguration);

    try {
      if (this.props.enableCache === false) {
        this.cache.remove(cacheConfiguration);
      } else if (cacheWasAllowedWhenLoadStarted) {
        const cachedApps = this.cache.read(cacheConfiguration, normalizeCacheDuration(this.props.cacheDurationMinutes));
        if (cachedApps !== undefined) {
          if (this.isCurrentLoad(loadRequestId)) {
            this.setState({
              apps: this.sortApps(cachedApps),
              error: undefined
            });
          }
          return;
        }
      }

      const { graphClient } = this.props;

      // App role assignments remain the source of truth for apps available to
      // the current user. Follow paging here as well because a user can have
      // more assignments than one Graph page contains.
      const assignments = await this.getAllGraphPages<IAppRoleAssignment>(
        graphClient
          .api('/me/appRoleAssignments')
          .select('id,principalDisplayName,resourceDisplayName,resourceId')
      );
      const myAssignedResourceIds = new Set(assignments.map(assignment => assignment.resourceId));

      // Only these service principals are Enterprise Application candidates.
      // In particular, ServiceIdentity (Entra Agent Identities) is excluded by
      // the servicePrincipalType filter rather than by a name-based heuristic.
      const integratedApps = await this.getAllGraphPages<IServicePrincipalInfo>(
        graphClient
          .api('/servicePrincipals')
          .filter("servicePrincipalType eq 'Application' and tags/any(t:t eq 'WindowsAzureActiveDirectoryIntegratedApp')")
          .select('id,appId,appOwnerOrganizationId,displayName,appDescription,notes,homepage,publisherName,verifiedPublisher,preferredSingleSignOnMode,info,tags,oauth2PermissionScopes')
          .top(999)
      );
      const includeDefaults = this.props.showDefaultApps ?? true;
      const visibleDefaultAppKeys = includeDefaults
        ? this.props.visibleDefaultAppNames.map(appName => this.normalizeName(appName))
        : [];
      const defaultAppKeys = defaultApps.map(defaultApp => this.normalizeName(defaultApp.name));

      const isSuppressedDefaultApp = (name: string): boolean => {
        const key = this.normalizeName(name);
        return defaultAppKeys.indexOf(key) !== -1 && visibleDefaultAppKeys.indexOf(key) === -1;
      };

      // Graph Enterprise Applications are keyed by servicePrincipal.id. This
      // prevents a candidate that is both assigned and integrated from being
      // added twice, while retaining name-based merging only for default apps.
      const enterpriseAppsById = new Map<string, IAppData>();
      const integratedAppsById = new Map<string, IServicePrincipalInfo>();
      integratedApps.forEach(app => integratedAppsById.set(app.id, app));

      assignments.forEach(assignment => {
        if (isSuppressedDefaultApp(assignment.resourceDisplayName)) {
          return;
        }

        const integratedApp = integratedAppsById.get(assignment.resourceId);
        enterpriseAppsById.set(assignment.resourceId, integratedApp
          ? this.createAppFromServicePrincipal(integratedApp)
          : this.createAppFromAssignment(assignment));
      });

      // Candidates which are already assigned do not need an assignment check.
      // For every other candidate, only a successful, explicitly empty response
      // qualifies it as unassigned; errors are intentionally fail-closed.
      const candidatesNeedingAssignmentCheck = integratedApps.filter(app =>
        !myAssignedResourceIds.has(app.id) && !isSuppressedDefaultApp(app.displayName || '')
      );
      const unassignedAppsResult = await this.getUnassignedIntegratedApps(candidatesNeedingAssignmentCheck);
      unassignedAppsResult.apps.forEach(app => {
        enterpriseAppsById.set(app.id, this.createAppFromServicePrincipal(app));
      });

      const standaloneDefaultApps: IAppData[] = [];

      if (includeDefaults) {
        defaultApps.forEach(defaultApp => {
          if (this.props.visibleDefaultAppNames.indexOf(defaultApp.name) === -1) {
            return;
          }

          const key = this.normalizeName(defaultApp.name);
          const existing = Array.from(enterpriseAppsById.values())
            .find(app => this.normalizeName(app.name) === key);

          if (existing) {
            enterpriseAppsById.set(existing.resourceId, {
              ...existing,
              name: defaultApp.name,
              url: defaultApp.url,
              iconUrl: defaultApp.icon,
              isDefaultApp: true
            });
          } else {
            standaloneDefaultApps.push({
              name: defaultApp.name,
              url: defaultApp.url,
              iconUrl: defaultApp.icon,
              resourceId: '',
              isHidden: false,
              isLoaded: true,
              isDefaultApp: true
            });
          }
        });
      }

      const allApps = Array.from(enterpriseAppsById.values()).concat(standaloneDefaultApps);
      if (allApps.length === 0) {
        if (this.isCurrentLoad(loadRequestId)) {
          this.setState({ apps: [] });
          if (unassignedAppsResult.isComplete && this.canWriteCache(cacheConfiguration, cacheWasAllowedWhenLoadStarted)) {
            this.cache.write(cacheConfiguration, []);
          }
        }
        return;
      }

      // Sort apps according to custom logic
      const appsArray = this.sortApps(allApps);

      // Set initial state with sorted apps
      if (this.isCurrentLoad(loadRequestId)) {
        this.setState({ apps: appsArray });
      }

      // Load service principal details asynchronously
      const loadedAppsResult = await this.loadServicePrincipalDetails(appsArray);

      if (this.isCurrentLoad(loadRequestId)) {
        this.setState({ apps: loadedAppsResult.apps });
        if (unassignedAppsResult.isComplete && loadedAppsResult.isComplete &&
            this.canWriteCache(cacheConfiguration, cacheWasAllowedWhenLoadStarted)) {
          this.cache.write(cacheConfiguration, loadedAppsResult.apps);
        }
      }

    } catch (error) {
      if (this.isCurrentLoad(loadRequestId)) {
        throw error;
      }
    }
    finally {
      if (this.isCurrentLoad(loadRequestId)) {
        this.setState({ isLoading: false }, this.openDetailForExactMatch);
      }
    }
  }

  private async getAllGraphPages<T>(initialRequest: IGraphRequest<T>): Promise<T[]> {
    const values: T[] = [];
    let response = await initialRequest.get();

    while (true) {
      values.push(...(response.value || []));
      const nextLink = response['@odata.nextLink'];
      if (!nextLink) {
        return values;
      }

      response = await this.props.graphClient.api(nextLink).get() as IGraphPage<T>;
    }
  }

  private createAppFromAssignment(assignment: IAppRoleAssignment): IAppData {
    return {
      name: assignment.resourceDisplayName,
      url: '',
      iconUrl: this.generateDefaultIcon(assignment.resourceDisplayName),
      resourceId: assignment.resourceId,
      isHidden: false,
      isLoaded: false,
      isDefaultApp: false
    };
  }

  private createAppFromServicePrincipal(servicePrincipal: IServicePrincipalInfo): IAppData {
    const name = servicePrincipal.displayName || servicePrincipal.appId;
    return this.applyServicePrincipalDetails({
      name,
      url: '',
      iconUrl: this.generateDefaultIcon(name),
      resourceId: servicePrincipal.id,
      isHidden: false,
      isLoaded: false,
      isDefaultApp: false
    }, servicePrincipal);
  }

  private chunk<T>(items: T[], size: number): T[][] {
    const chunks: T[][] = [];
    for (let index = 0; index < items.length; index += size) {
      chunks.push(items.slice(index, index + size));
    }
    return chunks;
  }

  private async getUnassignedIntegratedApps(candidates: IServicePrincipalInfo[]): Promise<IUnassignedAppsResult> {
    const unassignedApps: IServicePrincipalInfo[] = [];
    const batchSize = 20;
    let requestNumber = 0;
    let isComplete = true;

    for (const candidateChunk of this.chunk(candidates, batchSize)) {
      const requestIdToApp = new Map<string, IServicePrincipalInfo>();
      const requests = candidateChunk.map(servicePrincipal => {
        const id = `assignment-${requestNumber++}`;
        requestIdToApp.set(id, servicePrincipal);
        return {
          id,
          method: 'GET',
          // $top=1 is sufficient: a single assignment means that the app must
          // not be displayed, while an empty first page proves it is unassigned.
          url: `/servicePrincipals/${encodeURIComponent(servicePrincipal.id)}/appRoleAssignedTo?$select=principalId,principalType&$top=1`
        };
      });

      try {
        const batchResponse = await this.props.graphClient
          .api('/$batch')
          .post({ requests }) as IGraphBatchResponse;
        const responsesById = new Map<string, IGraphBatchSubResponse>();
        (batchResponse.responses || []).forEach(response => responsesById.set(response.id, response));

        requestIdToApp.forEach((servicePrincipal, requestId) => {
          const response = responsesById.get(requestId);
          const isSuccessful = response !== undefined && response.status >= 200 && response.status < 300;
          const assignments = response?.body?.value;

          if (!isSuccessful) {
            isComplete = false;
            console.warn(`Could not determine assignments for ${servicePrincipal.displayName || servicePrincipal.id}; app will not be shown.`, response);
          } else if (!Array.isArray(assignments)) {
            isComplete = false;
            console.warn(`Received an invalid assignment response for ${servicePrincipal.displayName || servicePrincipal.id}; app will not be shown.`, response);
          } else if (assignments.length === 0) {
            unassignedApps.push(servicePrincipal);
          }
        });
      } catch (error) {
        // A failed batch leaves the assignment state unknown, so all apps in
        // this chunk remain hidden rather than risking exposure to other users.
        isComplete = false;
        candidateChunk.forEach(servicePrincipal => {
          console.warn(`Could not determine assignments for ${servicePrincipal.displayName || servicePrincipal.id}; app will not be shown.`, error);
        });
      }
    }

    return { apps: unassignedApps, isComplete };
  }

  /**
   * Sort apps according to custom sort order and alphabetically
   */
  private sortApps(apps: IAppData[]): IAppData[] {
    const appsArray = apps.slice();

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
   * Load service principal details and prepare the final app state
   */
  private async loadServicePrincipalDetails(apps: IAppData[]): Promise<ILoadedAppsResult> {
    const { graphClient, showHiddenApps } = this.props;
    let isComplete = true;

    const loadPromises = apps.map(async (app) => {
      try {
        if (!app.resourceId) {
          return { ...app, isLoaded: true } as IAppData;
        }

        // The integrated-app query already returned these details. Reuse them
        // instead of issuing a redundant per-app request.
        if (app.servicePrincipal) {
          return this.applyServicePrincipalDetails(app, app.servicePrincipal);
        }

        const spInfo: IServicePrincipalInfo = await graphClient
          .api(`/servicePrincipals/${app.resourceId}`)
          .select('id,appId,appOwnerOrganizationId,displayName,appDescription,notes,homepage,publisherName,verifiedPublisher,preferredSingleSignOnMode,info,tags,oauth2PermissionScopes')
          .get();

        return this.applyServicePrincipalDetails(app, spInfo);
      } catch (error) {
        isComplete = false;
        console.warn(`Could not load details for ${app.name}:`, error);
        return { ...app, isLoaded: true };
      }
    });

    const results = await Promise.all(loadPromises);

    // Enterprise Applications were already deduplicated by servicePrincipal.id
    // before detail loading. Do not collapse distinct Graph resources merely
    // because they happen to share a display name.
    const dedupedApps = results;
    
    // Filter apps based on showHiddenApps setting
    const filteredApps = showHiddenApps 
      ? dedupedApps 
      : dedupedApps.filter(app => !app.isHidden);

    return { apps: filteredApps, isComplete };
  }

  private applyServicePrincipalDetails(app: IAppData, servicePrincipal: IServicePrincipalInfo): IAppData {
    const loginUrl = servicePrincipal.appOwnerOrganizationId
      ? `https://launcher.myapps.microsoft.com/api/signin/${servicePrincipal.appId}?tenantId=${servicePrincipal.appOwnerOrganizationId}`
      : app.url;
    const hasHideAppTag = Array.isArray(servicePrincipal.tags) && servicePrincipal.tags.indexOf('HideApp') !== -1;

    return {
      ...app,
      url: app.isDefaultApp ? app.url : loginUrl,
      iconUrl: servicePrincipal.info?.logoUrl || app.iconUrl,
      isHidden: hasHideAppTag,
      isLoaded: true,
      servicePrincipal
    };
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
      saml: strings.SamlSso,
      password: strings.PasswordSso,
      oidc: strings.OpenIdConnectSso,
      linked: strings.LinkedSso,
      notSupported: strings.SsoNotSupported
    };
    return labels[mode] || mode;
  }

  private renderOAuthScopes(servicePrincipal: IServicePrincipalInfo | undefined): React.ReactElement {
    const scopes = (servicePrincipal?.oauth2PermissionScopes || []).filter(scope => scope.isEnabled !== false);

    return (
      <section className={styles.detailInfoBlock}>
        <h3>{strings.OAuthScopes}</h3>
        {scopes.length === 0 ? (
          <p className={styles.notAvailable}>{strings.NotAvailable}</p>
        ) : (
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
        )}
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
    const showAppIdentifiers = this.props.displayAppIdentifiers;
    const showOAuthScopes = this.props.displayOAuthScopes;

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
