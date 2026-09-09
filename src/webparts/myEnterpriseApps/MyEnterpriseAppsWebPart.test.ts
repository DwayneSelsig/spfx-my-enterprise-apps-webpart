jest.mock('MyEnterpriseAppsWebPartStrings', () => ({}), { virtual: true });
jest.mock('@microsoft/sp-core-library', () => ({
  DisplayMode: { Read: 1, Edit: 2 },
  Version: { parse: jest.fn() }
}));
jest.mock('@microsoft/sp-property-pane', () => ({
  PropertyPaneCheckbox: jest.fn(),
  PropertyPaneDropdown: jest.fn(),
  PropertyPaneSlider: jest.fn(),
  PropertyPaneTextField: jest.fn(),
  PropertyPaneToggle: jest.fn()
}));
jest.mock('@microsoft/sp-webpart-base', () => ({
  BaseClientSideWebPart: class {
    public context = {};
    public displayMode = 1;
    public domElement = document.createElement('div');
    public properties = {};
  }
}));
jest.mock('./components/MyEnterpriseApps', () => () => null);

import type { MSGraphClientV3 } from '@microsoft/sp-http';
import MyEnterpriseAppsWebPart from './MyEnterpriseAppsWebPart';

interface IGraphClientAccessor {
  getGraphClient(): Promise<MSGraphClientV3>;
}

function createWebPart(getClient: jest.Mock): IGraphClientAccessor {
  const webPart = new MyEnterpriseAppsWebPart();
  Object.defineProperty(webPart, 'context', {
    configurable: true,
    value: {
      msGraphClientFactory: { getClient }
    }
  });

  return webPart as unknown as IGraphClientAccessor;
}

describe('MyEnterpriseAppsWebPart Graph client initialization', () => {
  it('shares a pending Graph client initialization between concurrent callers', async () => {
    const graphClient = {} as MSGraphClientV3;
    let resolveGraphClient!: (client: MSGraphClientV3) => void;
    const pendingClient = new Promise<MSGraphClientV3>(resolve => {
      resolveGraphClient = resolve;
    });
    const getClient = jest.fn(() => pendingClient);
    const webPart = createWebPart(getClient);

    const firstClient = webPart.getGraphClient();
    const secondClient = webPart.getGraphClient();

    expect(firstClient).toBe(secondClient);
    expect(getClient).toHaveBeenCalledTimes(1);
    expect(getClient).toHaveBeenCalledWith('3');

    resolveGraphClient(graphClient);

    await expect(firstClient).resolves.toBe(graphClient);
    await expect(secondClient).resolves.toBe(graphClient);
  });

  it('clears a failed initialization so a later call can succeed', async () => {
    const initializationError = new Error('Graph initialization failed');
    const graphClient = {} as MSGraphClientV3;
    const getClient = jest.fn()
      .mockRejectedValueOnce(initializationError)
      .mockResolvedValueOnce(graphClient);
    const webPart = createWebPart(getClient);

    await expect(webPart.getGraphClient()).rejects.toBe(initializationError);
    await expect(webPart.getGraphClient()).resolves.toBe(graphClient);

    expect(getClient).toHaveBeenCalledTimes(2);
    expect(getClient).toHaveBeenNthCalledWith(1, '3');
    expect(getClient).toHaveBeenNthCalledWith(2, '3');
  });
});
