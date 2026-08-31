import {
  DEFAULT_CACHE_DURATION_MINUTES,
  EnterpriseAppsCache,
  getEnterpriseAppsCacheKey,
  getEnterpriseAppsCacheSignature,
  normalizeCacheDuration,
  type IEnterpriseAppsCacheConfiguration
} from './EnterpriseAppsCache';
import type { IAppData } from './IMyEnterpriseAppsProps';

const configuration: IEnterpriseAppsCacheConfiguration = {
  tenantId: 'TENANT-ID',
  userId: 'USER-ID',
  showHiddenApps: false,
  showDefaultApps: true,
  visibleDefaultAppNames: ['Word', 'Excel']
};

const app: IAppData = {
  name: 'Contoso',
  url: 'https://contoso.example',
  iconUrl: 'data:image/svg+xml,test',
  resourceId: 'resource-id',
  isHidden: false,
  isLoaded: true,
  isDefaultApp: false
};

describe('EnterpriseAppsCache', () => {
  beforeEach(() => {
    window.localStorage.clear();
    jest.restoreAllMocks();
  });

  it('creates a tenant- and user-specific key', () => {
    expect(getEnterpriseAppsCacheKey(' TENANT-ID ', 'USER-ID')).toBe('myEnterpriseApps:tenant-id:user-id');
    expect(getEnterpriseAppsCacheKey('TENANT-ID', undefined)).toBeUndefined();
  });

  it('creates a stable signature for result-affecting configuration', () => {
    const reorderedConfiguration = {
      ...configuration,
      visibleDefaultAppNames: ['Excel', 'Word']
    };

    expect(getEnterpriseAppsCacheSignature(configuration)).toBe(
      getEnterpriseAppsCacheSignature(reorderedConfiguration)
    );
    expect(getEnterpriseAppsCacheSignature({ ...configuration, showHiddenApps: true }))
      .not.toBe(getEnterpriseAppsCacheSignature(configuration));
  });

  it('normalizes cache duration defensively', () => {
    expect(normalizeCacheDuration(undefined)).toBe(DEFAULT_CACHE_DURATION_MINUTES);
    expect(normalizeCacheDuration('invalid')).toBe(DEFAULT_CACHE_DURATION_MINUTES);
    expect(normalizeCacheDuration(1)).toBe(5);
    expect(normalizeCacheDuration(30)).toBe(30);
    expect(normalizeCacheDuration(1441)).toBe(1440);
  });

  it('writes and reads a valid cache entry', () => {
    const cache = new EnterpriseAppsCache();
    const cachedAt = 1_000_000;

    expect(cache.write(configuration, [app], cachedAt)).toBe(true);
    expect(cache.read(configuration, 30, cachedAt + 1)).toEqual([app]);

    const stored = JSON.parse(window.localStorage.getItem('myEnterpriseApps:tenant-id:user-id') as string) as Record<string, unknown>;
    expect(stored.schemaVersion).toBe(1);
    expect(stored.tenantId).toBe('tenant-id');
    expect(stored.userId).toBe('user-id');
    expect(stored.cachedAt).toBe(cachedAt);
  });

  it('expires entries at the configured TTL', () => {
    const cache = new EnterpriseAppsCache();
    const cachedAt = 1_000_000;

    cache.write(configuration, [app], cachedAt);

    expect(cache.read(configuration, 30, cachedAt + 30 * 60 * 1000 - 1)).toEqual([app]);
    expect(cache.read(configuration, 30, cachedAt + 30 * 60 * 1000)).toBeUndefined();
    expect(window.localStorage.getItem('myEnterpriseApps:tenant-id:user-id')).toBeNull();
  });

  it('does not reuse an entry for a different result configuration', () => {
    const cache = new EnterpriseAppsCache();

    cache.write(configuration, [app]);

    expect(cache.read({ ...configuration, showHiddenApps: true }, 30)).toBeUndefined();
    expect(window.localStorage.getItem('myEnterpriseApps:tenant-id:user-id')).toBeNull();
  });

  it('removes invalid JSON and incompatible entries', () => {
    const cache = new EnterpriseAppsCache();
    const key = 'myEnterpriseApps:tenant-id:user-id';
    const warn = jest.spyOn(console, 'warn').mockImplementation(() => undefined);

    window.localStorage.setItem(key, '{invalid');
    expect(cache.read(configuration, 30)).toBeUndefined();
    expect(window.localStorage.getItem(key)).toBeNull();

    window.localStorage.setItem(key, JSON.stringify({ schemaVersion: 0 }));
    expect(cache.read(configuration, 30)).toBeUndefined();
    expect(window.localStorage.getItem(key)).toBeNull();
    expect(warn).toHaveBeenCalled();
  });

  it('does not write incomplete app results', () => {
    const cache = new EnterpriseAppsCache();
    const incompleteApp = { ...app, isLoaded: false };

    expect(cache.write(configuration, [incompleteApp])).toBe(false);
    expect(window.localStorage.getItem('myEnterpriseApps:tenant-id:user-id')).toBeNull();
  });

  it('falls back safely when storage operations throw', () => {
    const cache = new EnterpriseAppsCache();
    const warn = jest.spyOn(console, 'warn').mockImplementation(() => undefined);

    const getItem = jest.spyOn(Storage.prototype, 'getItem').mockImplementation(() => {
      throw new Error('read failed');
    });
    expect(cache.read(configuration, 30)).toBeUndefined();

    getItem.mockRestore();
    jest.spyOn(Storage.prototype, 'setItem').mockImplementation(() => {
      throw new Error('quota exceeded');
    });
    expect(cache.write(configuration, [app])).toBe(false);
    expect(warn).toHaveBeenCalled();
  });
});
