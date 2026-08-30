import { MSGraphClientV3 } from '@microsoft/sp-http';

/**
 * Props for the MyEnterpriseApps React component
 */
export interface IMyEnterpriseAppsProps {
  title: string;
  sortOrder: string;
  showHiddenApps: boolean;
  showDefaultApps: boolean;
  displayInternalNotes: boolean;
  displayAppIdentifiers: boolean;
  displayOAuthScopes: boolean;
  enableDetailView: boolean;
  iconSize: number;
  textSize: number;
  appSpacing: number;
  hasTeamsContext: boolean;
  graphClient: MSGraphClientV3;
}

/**
 * State for the MyEnterpriseApps React component
 */
export interface IMyEnterpriseAppsState {
  apps: IAppData[];
  isLoading: boolean;
  error: string | undefined;
  filterQuery: string;
  isFilterOpen: boolean;
  selectedApp: IAppData | undefined;
  isDetailTransitioning: boolean;
  isReturningToResults: boolean;
  isDetailDismissed: boolean;
}

/**
 * Interface for Graph API app role assignments
 */
export interface IAppRoleAssignment {
  id: string;
  principalDisplayName: string;
  resourceDisplayName: string;
  resourceId: string;
}

/**
 * Interface for service principal information from Graph API
 */
export interface IServicePrincipalInfo {
  id: string;
  appId: string;
  appOwnerOrganizationId?: string;
  displayName?: string;
  appDescription?: string;
  notes?: string;
  homepage?: string;
  publisherName?: string;
  preferredSingleSignOnMode?: string;
  tags?: string[];
  verifiedPublisher?: {
    displayName?: string;
    verifiedPublisherId?: string;
  };
  info?: {
    logoUrl: string | undefined;
    termsOfServiceUrl?: string;
  };
  oauth2PermissionScopes?: IAppPermissionScope[];
}

export interface IAppPermissionScope {
  id: string;
  value: string;
  isEnabled?: boolean;
  adminConsentDisplayName?: string;
  adminConsentDescription?: string;
  userConsentDisplayName?: string;
  userConsentDescription?: string;
}

/**
 * Interface for app data used internally
 */
export interface IAppData {
  name: string;
  url: string;
  iconUrl: string;
  resourceId: string;
  isHidden: boolean;
  isLoaded: boolean;
  isDefaultApp?: boolean;
  servicePrincipal?: IServicePrincipalInfo;
}
