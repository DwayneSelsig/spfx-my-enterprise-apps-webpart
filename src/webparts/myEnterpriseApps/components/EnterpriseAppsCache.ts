import type { IAppData } from './IMyEnterpriseAppsProps';

export interface IEnterpriseAppsCacheConfiguration {
  tenantId?: string;
  userId?: string;
  showHiddenApps: boolean;
  showDefaultApps: boolean;
  visibleDefaultAppNames: string[];
}

export interface IEnterpriseAppsCacheEntry {
  schemaVersion: number;
  tenantId: string;
  userId: string;
  cachedAt: number;
  configSignature: string;
  apps: IAppData[];
}

export const ENTERPRISE_APPS_CACHE_SCHEMA_VERSION = 1;
export const DEFAULT_CACHE_DURATION_MINUTES = 30;
export const MIN_CACHE_DURATION_MINUTES = 10;
export const MAX_CACHE_DURATION_MINUTES = 600;
export const CACHE_DURATION_STEP_MINUTES = 10;

function normalizeIdentifier(identifier: unknown): string | undefined {
  const normalizedIdentifier = typeof identifier === 'string' ? identifier.trim().toLowerCase() : undefined;
  return normalizedIdentifier || undefined;
}

function normalizeVisibleDefaultAppNames(names: unknown): string[] {
  return Array.from(new Set((Array.isArray(names) ? names : [])
    .filter(name => typeof name === 'string')
    .map(name => name.trim().toLowerCase())
    .filter(Boolean)))
    .sort();
}

export function getEnterpriseAppsCacheKey(tenantId?: unknown, userId?: unknown): string | undefined {
  const normalizedTenantId = normalizeIdentifier(tenantId);
  const normalizedUserId = normalizeIdentifier(userId);

  if (!normalizedTenantId || !normalizedUserId) {
    return undefined;
  }

  return `myEnterpriseApps:${encodeURIComponent(normalizedTenantId)}:${encodeURIComponent(normalizedUserId)}`;
}

export function getEnterpriseAppsCacheSignature(configuration: IEnterpriseAppsCacheConfiguration): string {
  return JSON.stringify({
    showHiddenApps: configuration.showHiddenApps === true,
    showDefaultApps: configuration.showDefaultApps !== false,
    visibleDefaultAppNames: normalizeVisibleDefaultAppNames(configuration.visibleDefaultAppNames)
  });
}

export function normalizeCacheDuration(value: unknown): number {
  const numericValue = typeof value === 'number'
    ? value
    : typeof value === 'string' && value.trim() !== ''
      ? Number(value)
      : Number.NaN;

  if (!Number.isFinite(numericValue)) {
    return DEFAULT_CACHE_DURATION_MINUTES;
  }

  if (numericValue < MIN_CACHE_DURATION_MINUTES) {
    return MIN_CACHE_DURATION_MINUTES;
  }

  if (numericValue > MAX_CACHE_DURATION_MINUTES) {
    return MAX_CACHE_DURATION_MINUTES;
  }

  return Math.round(numericValue);
}

export class EnterpriseAppsCache {
  public read(
    configuration: IEnterpriseAppsCacheConfiguration,
    cacheDurationMinutes: unknown,
    now: number = Date.now()
  ): IAppData[] | undefined {
    const key = getEnterpriseAppsCacheKey(configuration.tenantId, configuration.userId);
    if (!key) {
      return undefined;
    }

    const storage = this.getStorage();
    if (!storage) {
      return undefined;
    }

    let rawValue: string | null;
    try {
      rawValue = storage.getItem(key);
    } catch (error) {
      this.warn('Could not read the browser cache; continuing without cached apps.', error);
      return undefined;
    }

    if (!rawValue) {
      return undefined;
    }

    let parsedValue: unknown;
    try {
      parsedValue = JSON.parse(rawValue) as unknown;
    } catch (error) {
      this.warn('The browser cache contained invalid JSON; removing it.', error);
      this.removeStorageItem(storage, key);
      return undefined;
    }

    if (!this.isValidCacheEntry(parsedValue, configuration)) {
      this.warn('The browser cache had an incompatible structure; removing it.');
      this.removeStorageItem(storage, key);
      return undefined;
    }

    const entry = parsedValue;
    const cacheAge = now - entry.cachedAt;
    const cacheDurationMilliseconds = normalizeCacheDuration(cacheDurationMinutes) * 60 * 1000;

    if (entry.cachedAt > now || cacheAge >= cacheDurationMilliseconds) {
      this.removeStorageItem(storage, key);
      return undefined;
    }

    return entry.apps;
  }

  public write(
    configuration: IEnterpriseAppsCacheConfiguration,
    apps: IAppData[],
    cachedAt: number = Date.now()
  ): boolean {
    const key = getEnterpriseAppsCacheKey(configuration.tenantId, configuration.userId);
    if (!key || !Number.isFinite(cachedAt) || cachedAt < 0 ||
        !Array.isArray(apps) || !apps.every(app => this.isValidApp(app) && app.isLoaded === true)) {
      if (key) {
        this.warn('The app result was incomplete; it was not written to the browser cache.');
      }
      return false;
    }

    const storage = this.getStorage();
    if (!storage) {
      return false;
    }

    const entry: IEnterpriseAppsCacheEntry = {
      schemaVersion: ENTERPRISE_APPS_CACHE_SCHEMA_VERSION,
      tenantId: normalizeIdentifier(configuration.tenantId) as string,
      userId: normalizeIdentifier(configuration.userId) as string,
      cachedAt,
      configSignature: getEnterpriseAppsCacheSignature(configuration),
      apps
    };

    let serializedEntry: string;
    try {
      serializedEntry = JSON.stringify(entry);
    } catch (error) {
      this.warn('Could not serialize the browser cache; continuing without caching.', error);
      return false;
    }

    try {
      storage.setItem(key, serializedEntry);
      return true;
    } catch (error) {
      this.warn('Could not write the browser cache, possibly because storage is full.', error);
      return false;
    }
  }

  public remove(configuration: IEnterpriseAppsCacheConfiguration): void {
    const key = getEnterpriseAppsCacheKey(configuration.tenantId, configuration.userId);
    if (!key) {
      return;
    }

    const storage = this.getStorage();
    if (storage) {
      this.removeStorageItem(storage, key);
    }
  }

  private getStorage(): Storage | undefined {
    try {
      return typeof window !== 'undefined' ? window.localStorage : undefined;
    } catch (error) {
      this.warn('Browser localStorage is unavailable; continuing without caching.', error);
      return undefined;
    }
  }

  private removeStorageItem(storage: Storage, key: string): void {
    try {
      storage.removeItem(key);
    } catch (error) {
      this.warn('Could not remove an invalid browser-cache entry.', error);
    }
  }

  private isValidCacheEntry(
    value: unknown,
    configuration: IEnterpriseAppsCacheConfiguration
  ): value is IEnterpriseAppsCacheEntry {
    if (!this.isRecord(value)) {
      return false;
    }

    const expectedTenantId = normalizeIdentifier(configuration.tenantId);
    const expectedUserId = normalizeIdentifier(configuration.userId);

    return value.schemaVersion === ENTERPRISE_APPS_CACHE_SCHEMA_VERSION &&
      typeof value.tenantId === 'string' && value.tenantId === expectedTenantId &&
      typeof value.userId === 'string' && value.userId === expectedUserId &&
      typeof value.cachedAt === 'number' && Number.isFinite(value.cachedAt) && value.cachedAt >= 0 &&
      value.configSignature === getEnterpriseAppsCacheSignature(configuration) &&
      Array.isArray(value.apps) && value.apps.every(app => this.isValidApp(app) && app.isLoaded === true);
  }

  private isValidApp(value: unknown): value is IAppData {
    if (!this.isRecord(value) ||
        typeof value.name !== 'string' ||
        typeof value.url !== 'string' ||
        typeof value.iconUrl !== 'string' ||
        typeof value.resourceId !== 'string' ||
        typeof value.isHidden !== 'boolean' ||
        typeof value.isLoaded !== 'boolean') {
      return false;
    }

    if (value.isDefaultApp !== undefined && typeof value.isDefaultApp !== 'boolean') {
      return false;
    }

    return value.servicePrincipal === undefined || value.servicePrincipal === null ||
      this.isValidServicePrincipal(value.servicePrincipal);
  }

  private isValidServicePrincipal(value: unknown): boolean {
    if (!this.isRecord(value) || typeof value.id !== 'string' || typeof value.appId !== 'string') {
      return false;
    }

    if (value.displayName !== undefined && typeof value.displayName !== 'string') {
      return false;
    }

    if (value.tags !== undefined && value.tags !== null &&
        (!Array.isArray(value.tags) || !value.tags.every(tag => typeof tag === 'string'))) {
      return false;
    }

    if (value.info !== undefined && value.info !== null &&
        (!this.isRecord(value.info) ||
          (value.info.logoUrl !== undefined && value.info.logoUrl !== null && typeof value.info.logoUrl !== 'string') ||
          (value.info.termsOfServiceUrl !== undefined && value.info.termsOfServiceUrl !== null && typeof value.info.termsOfServiceUrl !== 'string'))) {
      return false;
    }

    if (value.verifiedPublisher !== undefined && value.verifiedPublisher !== null &&
        (!this.isRecord(value.verifiedPublisher) ||
          (value.verifiedPublisher.displayName !== undefined && value.verifiedPublisher.displayName !== null && typeof value.verifiedPublisher.displayName !== 'string') ||
          (value.verifiedPublisher.verifiedPublisherId !== undefined && value.verifiedPublisher.verifiedPublisherId !== null && typeof value.verifiedPublisher.verifiedPublisherId !== 'string'))) {
      return false;
    }

    return value.oauth2PermissionScopes === undefined || value.oauth2PermissionScopes === null ||
      (Array.isArray(value.oauth2PermissionScopes) && value.oauth2PermissionScopes.every(scope =>
        this.isRecord(scope) && typeof scope.id === 'string' && typeof scope.value === 'string'));
  }

  private isRecord(value: unknown): value is Record<string, unknown> {
    return typeof value === 'object' && value !== undefined && value !== null;
  }

  private warn(message: string, error?: unknown): void {
    if (error === undefined) {
      console.warn(`[My Enterprise Apps] ${message}`);
    } else {
      console.warn(`[My Enterprise Apps] ${message}`, error);
    }
  }
}
