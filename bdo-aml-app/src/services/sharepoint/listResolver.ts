import { sharePointConfig } from '../../config/sharepointConfig';
import { columnSchema } from './columnSchema';
import { graphClient } from './graphClient';

export interface ResolvedListIds {
  clientOnboarding: string;
  directors: string;
  beneficialOwners: string;
  auditLogs: string;
  documentLibrary: string;
}

const GUID_PATTERN =
  /^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$/i;

function isValidGuid(value: string): boolean {
  return GUID_PATTERN.test(value.trim());
}

class SharePointListResolver {
  private siteListsCache: Map<string, { displayName: string; name: string; id: string }[]> =
    new Map();
  private resolvedIds: Partial<ResolvedListIds> = {};

  private async loadSiteLists(siteId: string) {
    if (this.siteListsCache.has(siteId)) {
      return this.siteListsCache.get(siteId)!;
    }
    const lists = await graphClient.getSiteLists(siteId);
    this.siteListsCache.set(siteId, lists);
    return lists;
  }

  private async resolveOne(
    siteId: string,
    key: keyof ResolvedListIds,
    configuredId: string | undefined,
    nameCandidates: readonly string[]
  ): Promise<string> {
    const cached = this.resolvedIds[key];
    if (cached) return cached;

    if (configuredId && isValidGuid(configuredId)) {
      try {
        await graphClient.getListMetadata(siteId, configuredId);
        this.resolvedIds[key] = configuredId;
        console.info(`[SharePoint] ${key}: using configured list ID ${configuredId}`);
        return configuredId;
      } catch {
        console.warn(
          `[SharePoint] Configured ID for "${key}" (${configuredId}) was not found — resolving by name…`
        );
      }
    } else if (configuredId) {
      console.warn(
        `[SharePoint] Invalid GUID for "${key}" (${configuredId}) — resolving by name…`
      );
    }

    const lists = await this.loadSiteLists(siteId);
    for (const candidate of nameCandidates) {
      const normalized = candidate.toLowerCase();
      const match = lists.find(
        (list) =>
          list.displayName.toLowerCase() === normalized ||
          list.name.toLowerCase() === normalized
      );
      if (match) {
        this.resolvedIds[key] = match.id;
        console.info(
          `[SharePoint] ${key}: resolved "${match.displayName}" → ${match.id}`
        );
        return match.id;
      }
    }

    const available = lists.map((l) => `${l.displayName} (${l.name})`).join(', ');
    throw new Error(
      `SharePoint list "${nameCandidates.join('" or "')}" was not found on site. ` +
        `Available lists: ${available || '(none)'}. ` +
        `Copy the correct List ID from SharePoint → List settings, or fix names in sharepointConfig.ts.`
    );
  }

  async resolveAll(siteId: string): Promise<ResolvedListIds> {
    const { lists, listNames, documentLibrary } = sharePointConfig;

    return {
      clientOnboarding: await this.resolveOne(
        siteId,
        'clientOnboarding',
        lists.clientOnboarding || undefined,
        listNames.clientOnboarding
      ),
      directors: await this.resolveOne(
        siteId,
        'directors',
        lists.directors || undefined,
        listNames.directors
      ),
      beneficialOwners: await this.resolveOne(
        siteId,
        'beneficialOwners',
        lists.beneficialOwners || undefined,
        listNames.beneficialOwners
      ),
      auditLogs: await this.resolveOne(
        siteId,
        'auditLogs',
        lists.auditLogs || undefined,
        listNames.auditLogs
      ),
      documentLibrary: await this.resolveOne(
        siteId,
        'documentLibrary',
        documentLibrary.listId || undefined,
        listNames.documentLibrary
      ),
    };
  }

  async clearCache() {
    this.siteListsCache.clear();
    this.resolvedIds = {};
    graphClient.clearListCache();
    columnSchema.clearCache();
  }
}

export const listResolver = new SharePointListResolver();
