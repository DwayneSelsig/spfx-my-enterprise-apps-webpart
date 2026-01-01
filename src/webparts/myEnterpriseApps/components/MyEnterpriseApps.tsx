import * as React from 'react';
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
  
  constructor(props: IMyEnterpriseAppsProps) {
    super(props);
    this.state = {
      apps: [],
      isLoading: true,
      error: undefined
    };
  }

  public componentDidMount(): void {
    this.loadApps().catch(error => {
      console.error('Error loading apps:', error);
      this.setState({ error: error.message, isLoading: false });
    });
  }

  public componentDidUpdate(prevProps: IMyEnterpriseAppsProps): void {
    // Reload if sort order or showHiddenApps changes
    if (prevProps.sortOrder !== this.props.sortOrder ||
        prevProps.showHiddenApps !== this.props.showHiddenApps ||
        prevProps.showDefaultApps !== this.props.showDefaultApps) {
      this.loadApps().catch(error => {
        console.error('Error loading apps:', error);
        this.setState({ error: error.message, isLoading: false });
      });
    }
  }

  /**
   * Get the CSS class for the current icon size
   */
  private getSizeClass(): string {
    const size = this.props.iconSize || 'normal';
    switch (size) {
      case 'small':
        return styles.sizeSmall;
      case 'large':
        return styles.sizeLarge;
      case 'huge':
        return styles.sizeHuge;
      case 'normal':
      default:
        return styles.sizeNormal;
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

      // Build a Map with all apps (keyed by normalized name)
      const allAppsMap = new Map<string, IAppData>();
      
      assignments.forEach(assignment => {
        const defaultIcon = this.generateDefaultIcon(assignment.resourceDisplayName);
        const key = this.normalizeName(assignment.resourceDisplayName);
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

      const includeDefaults = this.props.showDefaultApps ?? true;

      if (includeDefaults) {
        defaultApps.forEach(defaultApp => {
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
      this.setState({ isLoading: false });
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
          .select('appId,appOwnerOrganizationId,info,tags')
          .get();

        // Construct login URL
        const loginUrl = `https://launcher.myapps.microsoft.com/api/signin/${spInfo.appId}?tenantId=${spInfo.appOwnerOrganizationId}`;

        // Check if app has HideApp tag
        const hasHideAppTag = spInfo.tags && Array.isArray(spInfo.tags) && spInfo.tags.indexOf('HideApp') !== -1;

        // Update app data
        return {
          ...app,
          url: app.isDefaultApp ? app.url : loginUrl,
          iconUrl: spInfo.info?.logoUrl || app.iconUrl,
          isHidden: hasHideAppTag === true,
          isLoaded: true
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

    this.setState({ apps: filteredApps });
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
  private renderAppItem(app: IAppData): React.ReactElement {
    const appClasses = `${styles.appItem}${app.isHidden ? ` ${styles.hiddenApp}` : ''}`;

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

  public render(): React.ReactElement<IMyEnterpriseAppsProps> {
    const { title, hasTeamsContext } = this.props;
    const { apps, isLoading, error } = this.state;
    const displayTitle = title || strings.DefaultTitle;
    const sizeClass = this.getSizeClass();

    return (
      <section className={`${styles.myEnterpriseApps} ${sizeClass} ${hasTeamsContext ? styles.teams : ''}`}>
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
          
          <div className={styles.appsList}>
            {error && (
              <div className={styles.errorMessage}>
                {strings.ErrorLoading}: {error}
              </div>
            )}
            
            {isLoading && this.renderSkeletonItems()}
            
            {!isLoading && !error && apps.length === 0 && (
              <div className={styles.noAppsMessage}>
                {strings.NoAppsFound}
              </div>
            )}
            
            {!isLoading && !error && apps.map(app => this.renderAppItem(app))}
          </div>
        </div>
      </section>
    );
  }
}
