jest.mock('MyEnterpriseAppsWebPartStrings', () => ({
  DefaultTitle: 'My Apps', CopilotOpenLinkError: 'The link could not be opened.'
}), { virtual: true });
jest.mock('@microsoft/sp-copilot-component', () => ({
  BaseCopilotComponent: class {
    public context = {};
    public properties = {};
    public hostContext = {};
  }
}));
jest.mock('../../webparts/myEnterpriseApps/components/MyEnterpriseApps', () => {
  const react: typeof import('react') = jest.requireActual('react');
  return {
    __esModule: true,
    default: jest.fn(() => react.createElement('a', { href: 'https://example.com/app' }, 'Open app'))
  };
});

import { act } from 'react';
import type { MSGraphClientV3 } from '@microsoft/sp-http';
import MyEnterpriseApps from '../../webparts/myEnterpriseApps/components/MyEnterpriseApps';
import type { IMyEnterpriseAppsProps } from '../../webparts/myEnterpriseApps/components/IMyEnterpriseAppsProps';
import EnterpriseAppsCopilotComponent from './EnterpriseAppsCopilotComponent';
import propertiesSchema from './EnterpriseAppsCopilotComponentProperties';

interface ITestHost {
  context: {
    domElement: HTMLDivElement;
    msGraphClientFactory: { getClient: jest.Mock };
    pageContext: { aadInfo?: { tenantId: unknown; userId: unknown } };
    copilotBridge: { openLinkAsync: jest.Mock };
  };
  properties: { query?: string };
  hostContext: { theme?: string };
  render(): void;
  onDispose(): void;
}

const mockApps = MyEnterpriseApps as unknown as jest.Mock;
const tenantId = 'bf81e921-9518-4df7-b395-942a8cae976b';
const userId = '1fd5ef11-d4f0-4a07-a74d-b8c144bd36ab';
let host: ITestHost;
function renderedProps(): IMyEnterpriseAppsProps {
  return mockApps.mock.calls[mockApps.mock.calls.length - 1][0] as IMyEnterpriseAppsProps;
}
async function render(): Promise<void> {
  await act(async () => host.render());
}

beforeEach(() => {
  jest.clearAllMocks();
  const domElement = document.createElement('div');
  document.body.appendChild(domElement);
  host = new EnterpriseAppsCopilotComponent() as unknown as ITestHost;
  host.context = {
    domElement,
    msGraphClientFactory: { getClient: jest.fn() },
    pageContext: { aadInfo: { tenantId: { toString: () => tenantId }, userId: { toString: () => userId } } },
    copilotBridge: { openLinkAsync: jest.fn().mockResolvedValue({}) }
  };
  host.properties = { query: 'Salesforce' };
  host.hostContext = { theme: 'light' };
});
afterEach(() => {
  act(() => host.onDispose());
  host.context.domElement.remove();
});

it('passes the query and standard settings to the shared picker without acquiring Graph eagerly', async () => {
  await render();
  expect(renderedProps()).toMatchObject({ initialSearchQuery: 'Salesforce', showHiddenApps: false,
    showOnlyAssignedApps: true, showDefaultApps: true, enableCache: true, cacheDurationMinutes: 30,
    enableDetailView: true, isEditMode: false, isPropertyPaneOpen: false, tenantId, userId });
  expect(host.context.msGraphClientFactory.getClient).not.toHaveBeenCalled();
  expect(propertiesSchema.properties.query.type).toBe('string');
});

it('leaves cache identity unavailable when the host does not supply reliable identity', async () => {
  host.context.pageContext = {};
  await render();
  expect(renderedProps().tenantId).toBeUndefined();
  expect(renderedProps().userId).toBeUndefined();
  host.context.pageContext = { aadInfo: { tenantId: {}, userId: 'session-id' } };
  await render();
  expect(renderedProps().tenantId).toBeUndefined();
  expect(renderedProps().userId).toBeUndefined();
});

it('reuses a lazy Graph client and allows retry after client creation fails', async () => {
  await render();
  const client = {} as MSGraphClientV3;
  host.context.msGraphClientFactory.getClient.mockRejectedValueOnce(new Error('offline')).mockResolvedValue(client);
  await expect(renderedProps().getGraphClient()).rejects.toThrow('offline');
  const first = renderedProps().getGraphClient();
  expect(renderedProps().getGraphClient()).toBe(first);
  await expect(first).resolves.toBe(client);
  expect(host.context.msGraphClientFactory.getClient).toHaveBeenCalledTimes(2);
});

it('updates the theme and query through host rerendering', async () => {
  await render();
  host.hostContext = { theme: 'dark' };
  host.properties = { query: 'Teams' };
  await render();
  expect(renderedProps().initialSearchQuery).toBe('Teams');
  expect(renderedProps().bodyBackground).toBe('#1f1f1f');
});

it('opens links through the host bridge only after a user click and reports host rejection', async () => {
  await render();
  expect(host.context.copilotBridge.openLinkAsync).not.toHaveBeenCalled();
  host.context.copilotBridge.openLinkAsync.mockResolvedValue({ isError: true });
  await act(async () => {
    host.context.domElement.querySelector('a')!.dispatchEvent(new MouseEvent('click', { bubbles: true, cancelable: true }));
  });
  expect(host.context.copilotBridge.openLinkAsync).toHaveBeenCalledWith('https://example.com/app');
  expect(host.context.domElement.querySelector('[role="alert"]')?.textContent).toContain('could not be opened');
});
