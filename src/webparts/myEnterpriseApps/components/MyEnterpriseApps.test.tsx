/* The shared container is unmounted in afterEach after each helper render. */
/* eslint-disable @rushstack/pair-react-dom-render-unmount */

jest.mock('MyEnterpriseAppsWebPartStrings', () => ({
  DefaultTitle: 'My Apps',
  AllAppsLabel: 'Open My Apps',
  FilterAppsLabel: 'Filter apps',
  FilterAppsPlaceholder: 'Search apps',
  ClearFilterLabel: 'Clear filter',
  CloseFilterLabel: 'Close filter',
  NoFilterResults: 'No apps match your filter',
  NoAppsFound: 'No apps found',
  ErrorLoading: 'Error loading apps',
  BackToResults: 'Back to search results',
  OpenApp: 'Open app',
  NotAvailable: 'Not available',
  AppId: 'App ID',
  ServicePrincipalId: 'Service principal ID',
  Homepage: 'Homepage',
  Description: 'Description',
  SingleSignOn: 'Single sign-on',
  OpenTermsOfService: 'Open Terms of Service',
  OAuthScopes: 'Available OAuth scopes',
  Resources: 'Resources',
  VerifiedPublisher: 'Verified publisher',
  SamlSso: 'SAML',
  OpenIdConnectSso: 'OpenID Connect',
  PasswordSso: 'Password SSO',
  LinkedSso: 'Linked SSO',
  SsoNotSupported: 'Not supported'
}), { virtual: true });

import * as React from 'react';
import * as ReactDOM from 'react-dom';
import { act } from 'react-dom/test-utils';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { IMyEnterpriseAppsProps } from './IMyEnterpriseAppsProps';
import { EnterpriseAppsCache, type IEnterpriseAppsCacheConfiguration } from './EnterpriseAppsCache';
import MyEnterpriseApps from './MyEnterpriseApps';

interface IBatchRequest {
  id: string;
}

const cacheConfiguration: IEnterpriseAppsCacheConfiguration = {
  tenantId: 'tenant-id',
  userId: 'user-id',
  showHiddenApps: false,
  showDefaultApps: false,
  visibleDefaultAppNames: []
};

const cachedApp = {
  name: 'Cached app',
  url: 'https://cached.example',
  iconUrl: 'data:image/svg+xml,test',
  resourceId: 'cached-resource-id',
  isHidden: false,
  isLoaded: true,
  isDefaultApp: false
};

function createProps(graphClient: MSGraphClientV3): IMyEnterpriseAppsProps {
  return {
    title: 'Apps',
    sortOrder: '',
    showHiddenApps: false,
    showDefaultApps: false,
    visibleDefaultAppNames: [],
    enableCache: true,
    cacheDurationMinutes: 30,
    tenantId: 'tenant-id',
    userId: 'user-id',
    isPropertyPaneOpen: false,
    isEditMode: false,
    displayInternalNotes: false,
    displayAppIdentifiers: true,
    displayOAuthScopes: true,
    enableDetailView: true,
    iconSize: 48,
    textSize: 11,
    appSpacing: 12,
    hasTeamsContext: false,
    graphClient
  };
}

function createGraphClient(): { client: MSGraphClientV3; api: jest.Mock } {
  const assignment = {
    id: 'assignment-id',
    principalDisplayName: 'Contoso user',
    resourceDisplayName: 'Contoso',
    resourceId: 'resource-id'
  };
  const servicePrincipal = {
    id: 'resource-id',
    appId: 'app-id',
    appOwnerOrganizationId: 'tenant-id',
    displayName: 'Contoso',
    tags: ['WindowsAzureActiveDirectoryIntegratedApp'],
    info: {}
  };

  const api = jest.fn((path: string) => {
    const response = path === '/me/appRoleAssignments'
      ? { value: [assignment] }
      : { value: [servicePrincipal] };
    const request = {
      select: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
      get: jest.fn().mockResolvedValue(response)
    };
    return request;
  });

  return { client: { api } as unknown as MSGraphClientV3, api };
}

function createBatchGraphClient(appCount: number): { client: MSGraphClientV3; batchPosts: jest.Mock[] } {
  const integratedApps = Array.from({ length: appCount }, (_value, index) => ({
    id: `resource-id-${index}`,
    appId: `app-id-${index}`,
    appOwnerOrganizationId: 'tenant-id',
    displayName: `Contoso ${index}`,
    tags: ['WindowsAzureActiveDirectoryIntegratedApp'],
    info: {}
  }));
  const batchPosts: jest.Mock[] = [];
  const api = jest.fn((path: string) => {
    if (path === '/$batch') {
      const post = jest.fn((body: { requests: IBatchRequest[] }) => Promise.resolve({
        responses: body.requests.map(request => ({
          id: request.id,
          status: 200,
          body: { value: [] }
        }))
      }));
      batchPosts.push(post);
      return { post };
    }

    const response = path === '/me/appRoleAssignments'
      ? { value: [] }
      : { value: integratedApps };
    return {
      select: jest.fn().mockReturnThis(),
      filter: jest.fn().mockReturnThis(),
      top: jest.fn().mockReturnThis(),
      get: jest.fn().mockResolvedValue(response)
    };
  });

  return { client: { api } as unknown as MSGraphClientV3, batchPosts };
}

async function renderComponent(props: IMyEnterpriseAppsProps, container: HTMLElement): Promise<void> {
  await act(async () => {
    ReactDOM.render(React.createElement(MyEnterpriseApps, props), container);
    await new Promise<void>(resolve => window.setTimeout(resolve, 0));
  });
}

describe('MyEnterpriseApps cache integration', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    window.localStorage.clear();
    container = document.createElement('div');
    document.body.appendChild(container);
  });

  afterEach(() => {
    ReactDOM.unmountComponentAtNode(container);
    container.remove();
  });

  it('skips Graph when a valid cache entry exists', async () => {
    const cache = new EnterpriseAppsCache();
    cache.write(cacheConfiguration, [cachedApp]);
    const graphClient = { api: jest.fn() } as unknown as MSGraphClientV3;

    await renderComponent(createProps(graphClient), container);

    expect(graphClient.api).not.toHaveBeenCalled();
    expect(container.textContent).toContain('Cached app');

    const graphCallsBeforeSortChange = (graphClient.api as jest.Mock).mock.calls.length;
    await act(async () => {
      ReactDOM.render(
        React.createElement(MyEnterpriseApps, { ...createProps(graphClient), sortOrder: 'Cached' }),
        container
      );
      await new Promise<void>(resolve => window.setTimeout(resolve, 0));
    });
    expect((graphClient.api as jest.Mock).mock.calls.length).toBe(graphCallsBeforeSortChange);
  });

  it('writes only the complete result after a Graph load', async () => {
    const graphClient = createGraphClient();

    await renderComponent(createProps(graphClient.client), container);

    expect(graphClient.api).toHaveBeenCalledWith('/me/appRoleAssignments');
    expect(graphClient.api).toHaveBeenCalledWith('/servicePrincipals');
    expect(container.textContent).toContain('Contoso');
    expect(new EnterpriseAppsCache().read(cacheConfiguration, 30)).toHaveLength(1);
  });

  it('bypasses the cache while the Property Pane is open and reuses it after closing', async () => {
    const cache = new EnterpriseAppsCache();
    cache.write(cacheConfiguration, [cachedApp]);
    const graphClient = createGraphClient();

    await renderComponent({ ...createProps(graphClient.client), isPropertyPaneOpen: true }, container);

    expect(graphClient.api).toHaveBeenCalledWith('/me/appRoleAssignments');
    expect(new EnterpriseAppsCache().read(cacheConfiguration, 30)).toEqual([cachedApp]);
    const graphCallsWhileEditing = graphClient.api.mock.calls.length;

    await act(async () => {
      ReactDOM.render(
        React.createElement(MyEnterpriseApps, { ...createProps(graphClient.client), isPropertyPaneOpen: false }),
        container
      );
      await new Promise<void>(resolve => window.setTimeout(resolve, 0));
    });

    expect(graphClient.api.mock.calls.length).toBe(graphCallsWhileEditing);
    expect(container.textContent).toContain('Cached app');
  });

  it('removes and bypasses the cache when caching is disabled', async () => {
    const cache = new EnterpriseAppsCache();
    cache.write(cacheConfiguration, [cachedApp]);
    const graphClient = createGraphClient();

    await renderComponent({ ...createProps(graphClient.client), enableCache: false }, container);

    expect(graphClient.api).toHaveBeenCalledWith('/me/appRoleAssignments');
    expect(new EnterpriseAppsCache().read(cacheConfiguration, 30)).toBeUndefined();
  });

  it.each([20, 21])('keeps Graph batches at a maximum of 20 requests for %i apps', async (appCount: number) => {
    const graphClient = createBatchGraphClient(appCount);

    await renderComponent({ ...createProps(graphClient.client), enableCache: false }, container);

    expect(graphClient.batchPosts).toHaveLength(appCount === 20 ? 1 : 2);
    expect((graphClient.batchPosts[0].mock.calls[0][0] as { requests: IBatchRequest[] }).requests).toHaveLength(20);
    if (appCount === 21) {
      expect((graphClient.batchPosts[1].mock.calls[0][0] as { requests: IBatchRequest[] }).requests).toHaveLength(1);
    }
  });
});
