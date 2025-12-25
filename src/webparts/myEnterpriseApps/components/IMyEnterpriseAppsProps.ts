import { MSGraphClientV3 } from '@microsoft/sp-http';

/**
 * Props for the MyEnterpriseApps React component
 */
export interface IMyEnterpriseAppsProps {
  title: string;
  sortOrder: string;
  showHiddenApps: boolean;
  iconSize: 'small' | 'normal' | 'large' | 'huge';
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
  appId: string;
  appOwnerOrganizationId: string;
  tags?: string[];
  info: {
    logoUrl: string | undefined;
  };
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
}
