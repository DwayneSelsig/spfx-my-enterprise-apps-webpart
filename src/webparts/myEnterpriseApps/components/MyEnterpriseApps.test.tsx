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
import { act } from 'react';
import { initializeIcons } from '@fluentui/react/lib/Icons';
import { createRoot, type Root } from 'react-dom/client';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import type { IMyEnterpriseAppsProps } from './IMyEnterpriseAppsProps';
import { EnterpriseAppsCache, type IEnterpriseAppsCacheConfiguration } from './EnterpriseAppsCache';
import MyEnterpriseApps from './MyEnterpriseApps';

initializeIcons();

interface IBatchRequest {
  id: string;
}

const cacheConfiguration: IEnterpriseAppsCacheConfiguration = {
  tenantId: 'tenant-id',
  userId: 'user-id',
  showHiddenApps: false,
  showOnlyAssignedApps: false,
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

function createProps(
  graphClient: MSGraphClientV3,
  getGraphClient: () => Promise<MSGraphClientV3> = jest.fn().mockResolvedValue(graphClient)
): IMyEnterpriseAppsProps {
  return {
    title: 'Apps',
    sortOrder: '',
    showHiddenApps: false,
    showOnlyAssignedApps: false,
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
    themePrimary: '#0078d4',
    bodyBackground: '#ffffff',
    themePrimaryTextColor: '#ffffff',
    hasTeamsContext: false,
    getGraphClient
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
      : path === '/servicePrincipals'
        ? { value: [servicePrincipal] }
        : servicePrincipal;
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

function createBatchGraphClient(
  appCount: number,
  hasAssignments: boolean = false
): { client: MSGraphClientV3; batchPosts: jest.Mock[] } {
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
          body: {
            value: hasAssignments
              ? [{ principalId: 'assigned-principal-id', principalType: 'User' }]
              : []
          }
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

let root: Root;

async function renderComponent(props: IMyEnterpriseAppsProps): Promise<void> {
  await act(async () => {
    root.render(React.createElement(MyEnterpriseApps, props));
    await new Promise<void>(resolve => window.setTimeout(resolve, 0));
  });
}

async function flushMicrotasks(): Promise<void> {
  for (let index = 0; index < 10; index++) {
    await Promise.resolve();
  }
}

describe('MyEnterpriseApps cache integration', () => {
  let container: HTMLDivElement;

  beforeEach(() => {
    window.localStorage.clear();
    container = document.createElement('div');
    document.body.appendChild(container);
    root = createRoot(container);
  });

  afterEach(() => {
    act(() => root.unmount());
    container.remove();
    jest.restoreAllMocks();
  });

  it.each(['', 'cached', 'CaChEd', 'missing'])(
    'applies the host search %s using the existing app-name filter', async query => {
      new EnterpriseAppsCache().write(cacheConfiguration, [cachedApp, {
        ...cachedApp, name: 'Other app', resourceId: 'other-resource-id'
      }]);
      const graphClient = { api: jest.fn() } as unknown as MSGraphClientV3;
      await renderComponent({ ...createProps(graphClient), enableDetailView: false, initialSearchQuery: query });

      if (!query) {
        expect(container.textContent).toContain('Cached app');
        expect(container.textContent).toContain('Other app');
      } else if (query === 'missing') {
        expect(container.textContent).toContain('No apps match your filter');
      } else {
        expect(container.textContent).toContain('Cached app');
        expect(container.textContent).not.toContain('Other app');
      }
      expect(graphClient.api).not.toHaveBeenCalled();
    }
  );

  it('opens the existing detail view when the initial search matches one loaded app', async () => {
    new EnterpriseAppsCache().write(cacheConfiguration, [cachedApp]);
    const graphClient = { api: jest.fn() } as unknown as MSGraphClientV3;
    await renderComponent({ ...createProps(graphClient), initialSearchQuery: 'cached' });
    await act(async () => { await new Promise<void>(resolve => window.setTimeout(resolve, 200)); });

    expect(container.textContent).toContain('Back to search results');
    const openApp = container.querySelector<HTMLAnchorElement>('a[href="https://cached.example"]');
    expect(openApp?.textContent).toContain('Open app');
    expect(graphClient.api).not.toHaveBeenCalled();
  });

  it('keeps a search received before Graph finishes loading and applies the normal detail workflow', async () => {
    const graphClient = createGraphClient();
    let finishClient: (client: MSGraphClientV3) => void = () => undefined;
    const pendingClient = new Promise<MSGraphClientV3>(resolve => { finishClient = resolve; });
    const getGraphClient = jest.fn().mockReturnValue(pendingClient);
    const props = { ...createProps(graphClient.client, getGraphClient), initialSearchQuery: 'absent' };
    await renderComponent(props);
    await renderComponent({ ...props, initialSearchQuery: 'contoso' });
    expect(container.querySelector<HTMLInputElement>('input')?.value).toBe('contoso');
    await act(async () => { finishClient(graphClient.client); await flushMicrotasks(); });
    await act(async () => { await new Promise<void>(resolve => window.setTimeout(resolve, 200)); });

    expect(container.textContent).toContain('Back to search results');
    expect(container.textContent).toContain('Contoso');
    expect(getGraphClient).toHaveBeenCalledTimes(1);
  });

  it('accepts a new host search but preserves manual edits during ordinary rerenders', async () => {
    new EnterpriseAppsCache().write(cacheConfiguration, [cachedApp]);
    const graphClient = { api: jest.fn() } as unknown as MSGraphClientV3;
    const props = { ...createProps(graphClient), enableDetailView: false, initialSearchQuery: 'cached' };
    await renderComponent(props);
    const input = container.querySelector<HTMLInputElement>('input')!;
    await act(async () => {
      const setter = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, 'value')!.set!;
      setter.call(input, 'manual');
      input.dispatchEvent(new Event('input', { bubbles: true }));
    });
    expect(input.value).toBe('manual');
    await renderComponent({ ...props, bodyBackground: '#111111' });
    expect(container.querySelector<HTMLInputElement>('input')?.value).toBe('manual');
    await renderComponent({ ...props, initialSearchQuery: 'next' });
    expect(container.querySelector<HTMLInputElement>('input')?.value).toBe('next');
    await renderComponent({ ...props, initialSearchQuery: '' });
    expect(container.querySelector('input')).toBeNull();
    expect(container.textContent).toContain('Cached app');
    expect(graphClient.api).not.toHaveBeenCalled();
  });

  it('skips Graph when a valid cache entry exists', async () => {
    const cache = new EnterpriseAppsCache();
    cache.write(cacheConfiguration, [cachedApp]);
    const graphClient = { api: jest.fn() } as unknown as MSGraphClientV3;
    const getGraphClient = jest.fn().mockResolvedValue(graphClient);

    await renderComponent(createProps(graphClient, getGraphClient));

    expect(getGraphClient).not.toHaveBeenCalled();
    expect(graphClient.api).not.toHaveBeenCalled();
    expect(container.textContent).toContain('Cached app');

    const graphCallsBeforeSortChange = (graphClient.api as jest.Mock).mock.calls.length;
    await act(async () => {
      root.render(
        React.createElement(MyEnterpriseApps, { ...createProps(graphClient, getGraphClient), sortOrder: 'Cached' }));
      await new Promise<void>(resolve => window.setTimeout(resolve, 0));
    });
    expect(getGraphClient).not.toHaveBeenCalled();
    expect((graphClient.api as jest.Mock).mock.calls.length).toBe(graphCallsBeforeSortChange);
  });

  it('initializes Graph lazily on a cache miss and writes only the complete result', async () => {
    const graphClient = createGraphClient();
    const getGraphClient = jest.fn().mockResolvedValue(graphClient.client);

    await renderComponent(createProps(graphClient.client, getGraphClient));

    expect(getGraphClient).toHaveBeenCalledTimes(1);
    expect(graphClient.api).toHaveBeenCalledWith('/me/appRoleAssignments');
    expect(graphClient.api).toHaveBeenCalledWith('/servicePrincipals');
    expect(container.textContent).toContain('Contoso');
    expect(new EnterpriseAppsCache().read(cacheConfiguration, 30)).toHaveLength(1);
  });

  it('initializes Graph lazily when the cache entry has expired', async () => {
    const cache = new EnterpriseAppsCache();
    cache.write(cacheConfiguration, [cachedApp], Date.now() - (31 * 60 * 1000));
    const graphClient = createGraphClient();
    const getGraphClient = jest.fn().mockResolvedValue(graphClient.client);

    await renderComponent(createProps(graphClient.client, getGraphClient));

    expect(getGraphClient).toHaveBeenCalledTimes(1);
    expect(graphClient.api).toHaveBeenCalledWith('/me/appRoleAssignments');
    expect(container.textContent).toContain('Contoso');
    expect(container.textContent).not.toContain('Cached app');
    expect(cache.read(cacheConfiguration, 30)).toHaveLength(1);
  });

  it('waits for Graph initialization after an expired cache entry before calling api', async () => {
    const cache = new EnterpriseAppsCache();
    cache.write(cacheConfiguration, [cachedApp], Date.now() - (31 * 60 * 1000));
    const graphClient = createGraphClient();
    let resolveGraphClient!: (client: MSGraphClientV3) => void;
    const graphClientPromise = new Promise<MSGraphClientV3>(resolve => {
      resolveGraphClient = resolve;
    });
    const getGraphClient = jest.fn(() => graphClientPromise);
    const error = jest.spyOn(console, 'error').mockImplementation(() => undefined);

    act(() => {
      root.render(
        React.createElement(MyEnterpriseApps, createProps(graphClient.client, getGraphClient)));
    });

    expect(getGraphClient).toHaveBeenCalledTimes(1);
    expect(graphClient.api).not.toHaveBeenCalled();
    expect(error).not.toHaveBeenCalled();
    expect(container.querySelectorAll('[class*="skeleton"]')).not.toHaveLength(0);

    await act(async () => {
      resolveGraphClient(graphClient.client);
      await flushMicrotasks();
    });

    expect(graphClient.api).toHaveBeenCalledWith('/me/appRoleAssignments');
    expect(container.textContent).toContain('Contoso');
    expect(container.querySelectorAll('[class*="skeleton"]')).toHaveLength(0);
    expect(cache.read(cacheConfiguration, 30)).toHaveLength(1);
  });

  it('does not start Graph calls for a load that becomes stale while awaiting the client', async () => {
    const staleGraphClient = { api: jest.fn() } as unknown as MSGraphClientV3;
    const currentGraphClient = createGraphClient();
    let resolveStaleGraphClient!: (client: MSGraphClientV3) => void;
    const staleGraphClientPromise = new Promise<MSGraphClientV3>(resolve => {
      resolveStaleGraphClient = resolve;
    });
    const getGraphClient = jest.fn()
      .mockImplementationOnce(() => staleGraphClientPromise)
      .mockResolvedValue(currentGraphClient.client);

    act(() => {
      root.render(
        React.createElement(MyEnterpriseApps, createProps(staleGraphClient, getGraphClient)));
    });

    await act(async () => {
      root.render(
        React.createElement(MyEnterpriseApps, {
          ...createProps(currentGraphClient.client, getGraphClient),
          showOnlyAssignedApps: true
        }));
      await flushMicrotasks();
    });

    await act(async () => {
      resolveStaleGraphClient(staleGraphClient);
      await flushMicrotasks();
    });

    expect(getGraphClient).toHaveBeenCalledTimes(2);
    expect(staleGraphClient.api).not.toHaveBeenCalled();
    expect(currentGraphClient.api).toHaveBeenCalledWith('/me/appRoleAssignments');
  });

  it('routes Graph client initialization failures through the existing load error handling', async () => {
    const graphClient = { api: jest.fn() } as unknown as MSGraphClientV3;
    const getGraphClient = jest.fn().mockRejectedValue(new Error('Graph initialization failed'));
    jest.spyOn(console, 'error').mockImplementation(() => undefined);

    await renderComponent(createProps(graphClient, getGraphClient));

    expect(getGraphClient).toHaveBeenCalledTimes(1);
    expect(graphClient.api).not.toHaveBeenCalled();
    expect(container.textContent).toContain('Error loading apps: Graph initialization failed');
    expect(container.querySelectorAll('[class*="skeleton"]')).toHaveLength(0);
  });

  it('skips tenant-wide and unassigned-app queries when only assigned apps are shown', async () => {
    const graphClient = createGraphClient();
    const assignedOnlyConfiguration = { ...cacheConfiguration, showOnlyAssignedApps: true };

    await renderComponent({ ...createProps(graphClient.client), showOnlyAssignedApps: true });

    expect(graphClient.api).toHaveBeenCalledWith('/me/appRoleAssignments');
    expect(graphClient.api).not.toHaveBeenCalledWith('/servicePrincipals');
    expect(graphClient.api).toHaveBeenCalledWith('/servicePrincipals/resource-id');
    expect(container.textContent).toContain('Contoso');
    expect(new EnterpriseAppsCache().read(assignedOnlyConfiguration, 30)).toHaveLength(1);
  });

  it('bypasses the cache while the Property Pane is open and reuses it after closing', async () => {
    const cache = new EnterpriseAppsCache();
    cache.write(cacheConfiguration, [cachedApp]);
    const graphClient = createGraphClient();

    await renderComponent({ ...createProps(graphClient.client), isPropertyPaneOpen: true });

    expect(graphClient.api).toHaveBeenCalledWith('/me/appRoleAssignments');
    expect(new EnterpriseAppsCache().read(cacheConfiguration, 30)).toEqual([cachedApp]);
    const graphCallsWhileEditing = graphClient.api.mock.calls.length;

    await act(async () => {
      root.render(
        React.createElement(MyEnterpriseApps, { ...createProps(graphClient.client), isPropertyPaneOpen: false }));
      await new Promise<void>(resolve => window.setTimeout(resolve, 0));
    });

    expect(graphClient.api.mock.calls.length).toBe(graphCallsWhileEditing);
    expect(container.textContent).toContain('Cached app');
  });

  it('removes and bypasses the cache when caching is disabled', async () => {
    const cache = new EnterpriseAppsCache();
    cache.write(cacheConfiguration, [cachedApp]);
    const graphClient = createGraphClient();

    await renderComponent({ ...createProps(graphClient.client), enableCache: false });

    expect(graphClient.api).toHaveBeenCalledWith('/me/appRoleAssignments');
    expect(new EnterpriseAppsCache().read(cacheConfiguration, 30)).toBeUndefined();
  });

  it('writes the cache after a successful assignment check finds existing assignments', async () => {
    const graphClient = createBatchGraphClient(1, true);
    const warn = jest.spyOn(console, 'warn').mockImplementation(() => undefined);

    await renderComponent(createProps(graphClient.client));

    expect(new EnterpriseAppsCache().read(cacheConfiguration, 30)).toEqual([]);
    expect(warn.mock.calls.some(([message]) =>
      typeof message === 'string' && message.indexOf('assignment response') !== -1
    )).toBe(false);
  });

  it.each([20, 21])('keeps Graph batches at a maximum of 20 requests for %i apps', async (appCount: number) => {
    const graphClient = createBatchGraphClient(appCount);

    await renderComponent({ ...createProps(graphClient.client), enableCache: false });

    expect(graphClient.batchPosts).toHaveLength(appCount === 20 ? 1 : 2);
    expect((graphClient.batchPosts[0].mock.calls[0][0] as { requests: IBatchRequest[] }).requests).toHaveLength(20);
    if (appCount === 21) {
      expect((graphClient.batchPosts[1].mock.calls[0][0] as { requests: IBatchRequest[] }).requests).toHaveLength(1);
    }
  });
});
