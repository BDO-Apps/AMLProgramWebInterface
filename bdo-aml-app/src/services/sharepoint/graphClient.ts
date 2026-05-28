import { msalService } from '../MSALService';

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

export interface GraphListItem {
  id: string;
  fields: Record<string, unknown>;
}

interface GraphListResponse {
  value: GraphListItem[];
  '@odata.nextLink'?: string;
}

export class GraphClient {
  private siteId: string | null = null;
  private driveId: string | null = null;
  private driveListId: string | null = null;

  private async getToken(): Promise<string> {
    return msalService.getAccessToken();
  }

  private async request<T>(
    path: string,
    options: RequestInit = {}
  ): Promise<T> {
    const token = await this.getToken();
    const response = await fetch(`${GRAPH_BASE}${path}`, {
      ...options,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
        Prefer: 'HonorNonIndexedQueriesWarningMayFailRandomly',
        ...(options.headers as Record<string, string>),
      },
    });

    if (!response.ok) {
      const body = await response.text();
      let detail = body;
      try {
        const json = JSON.parse(body) as { error?: { message?: string } };
        detail = json.error?.message ?? body;
      } catch {
        /* use raw body */
      }
      throw new Error(
        `Microsoft Graph error (${response.status}) on ${path}: ${detail}`
      );
    }

    if (response.status === 204) {
      return undefined as T;
    }

    return response.json() as Promise<T>;
  }

  async getSiteId(hostname: string, sitePath: string): Promise<string> {
    if (this.siteId) return this.siteId;
    const data = await this.request<{ id: string }>(
      `/sites/${hostname}:${sitePath}`
    );
    this.siteId = data.id;
    return data.id;
  }

  clearListCache() {
    this.driveId = null;
    this.driveListId = null;
  }

  async getSiteLists(
    siteId: string
  ): Promise<{ id: string; displayName: string; name: string }[]> {
    const data = await this.request<{
      value: { id: string; displayName: string; name: string }[];
    }>(`/sites/${siteId}/lists?$select=id,displayName,name`);
    return data.value;
  }

  async getListMetadata(siteId: string, listId: string): Promise<{ id: string }> {
    return this.request<{ id: string }>(`/sites/${siteId}/lists/${listId}?$select=id`);
  }

  async getDriveId(siteId: string, libraryListId: string): Promise<string> {
    if (this.driveId && this.driveListId === libraryListId) return this.driveId;
    const data = await this.request<{ id: string }>(
      `/sites/${siteId}/lists/${libraryListId}/drive`
    );
    this.driveId = data.id;
    this.driveListId = libraryListId;
    return data.id;
  }

  async listAllItems(siteId: string, listId: string): Promise<GraphListItem[]> {
    const items: GraphListItem[] = [];
    let path: string | null =
      `/sites/${siteId}/lists/${listId}/items?expand=fields&$top=200`;

    while (path) {
      const page: GraphListResponse = await this.request<GraphListResponse>(path);
      items.push(...page.value);
      const nextLink = page['@odata.nextLink'];
      path = nextLink ? nextLink.replace(GRAPH_BASE, '') : null;
    }

    return items;
  }

  async getItemByTitle(
    siteId: string,
    listId: string,
    title: string
  ): Promise<GraphListItem | null> {
    const escaped = title.replace(/'/g, "''");
    const data = await this.request<GraphListResponse>(
      `/sites/${siteId}/lists/${listId}/items?$expand=fields&$filter=fields/Title eq '${escaped}'&$top=1`
    );
    return data.value[0] ?? null;
  }

  async createListItem(
    siteId: string,
    listId: string,
    fields: Record<string, unknown>
  ): Promise<GraphListItem> {
    return this.request<GraphListItem>(`/sites/${siteId}/lists/${listId}/items`, {
      method: 'POST',
      body: JSON.stringify({ fields }),
    });
  }

  async updateListItem(
    siteId: string,
    listId: string,
    itemId: string,
    fields: Record<string, unknown>
  ): Promise<void> {
    await this.request(
      `/sites/${siteId}/lists/${listId}/items/${itemId}/fields`,
      {
        method: 'PATCH',
        body: JSON.stringify(fields),
      }
    );
  }

  async deleteListItem(
    siteId: string,
    listId: string,
    itemId: string
  ): Promise<void> {
    await this.request(`/sites/${siteId}/lists/${listId}/items/${itemId}`, {
      method: 'DELETE',
    });
  }

  async uploadFileToFolder(
    siteId: string,
    driveId: string,
    folderPath: string,
    fileName: string,
    file: File,
    metadata: Record<string, unknown>
  ): Promise<{ id: string; webUrl?: string }> {
    const token = await this.getToken();
    const encodedPath = [folderPath, fileName]
      .map((segment) => encodeURIComponent(segment))
      .join('/');

    const uploadUrl = `${GRAPH_BASE}/sites/${siteId}/drives/${driveId}/root:/${encodedPath}:/content`;

    const response = await fetch(uploadUrl, {
      method: 'PUT',
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': file.type || 'application/octet-stream',
      },
      body: file,
    });

    if (!response.ok) {
      const body = await response.text();
      throw new Error(`File upload failed (${response.status}): ${body}`);
    }

    const driveItem = (await response.json()) as { id: string; webUrl?: string };

    if (driveItem.id && Object.keys(metadata).length > 0) {
      await this.request(
        `/sites/${siteId}/drives/${driveId}/items/${driveItem.id}/listItem/fields`,
        {
          method: 'PATCH',
          body: JSON.stringify(metadata),
        }
      );
    }

    return driveItem;
  }

  async listFilesInFolder(
    siteId: string,
    driveId: string,
    folderPath: string
  ): Promise<
    Array<{
      id: string;
      name: string;
      size: number;
      createdDateTime: string;
      webUrl?: string;
      fields: Record<string, unknown>;
    }>
  > {
    try {
      const encoded = encodeURIComponent(folderPath);
      const data = await this.request<{
        value: Array<{
          id: string;
          name: string;
          size: number;
          createdDateTime: string;
          webUrl?: string;
          file?: object;
          listItem?: { fields?: Record<string, unknown> };
        }>;
      }>(`/sites/${siteId}/drives/${driveId}/root:/${encoded}:/children?$expand=listItem`);

      return data.value
        .filter((item) => item.file)
        .map((item) => ({
          id: item.id,
          name: item.name,
          size: item.size,
          createdDateTime: item.createdDateTime,
          webUrl: item.webUrl,
          fields: item.listItem?.fields ?? {},
        }));
    } catch {
      return [];
    }
  }

  async getListColumns(
    siteId: string,
    listId: string
  ): Promise<Array<{ name: string; displayName: string; type: string; readOnly: boolean }>> {
    const data = await this.request<{
      value: Array<
        Record<string, unknown> & {
          name: string;
          displayName: string;
          readOnly?: boolean;
        }
      >;
    }>(`/sites/${siteId}/lists/${listId}/columns`);

    return data.value.map((col) => ({
      name: col.name,
      displayName: col.displayName,
      type: inferColumnType(col),
      readOnly: col.readOnly === true,
    }));
  }
}

function inferColumnType(col: Record<string, unknown>): string {
  if (col.text) return 'text';
  if (col.boolean) return 'boolean';
  if (col.choice) return 'choice';
  if (col.dateTime) return 'dateTime';
  if (col.lookup) return 'lookup';
  if (col.personOrGroup) return 'person';
  if (col.number) return 'number';
  return 'unknown';
}

export const graphClient = new GraphClient();
