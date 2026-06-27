import { SPHttpClient, SPHttpClientResponse } from '@microsoft/sp-http';
import { HttpClient, HttpClientResponse } from '@microsoft/sp-http';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { parseSpError } from '../common/helpers/errorHelper';
import { getConfig, IEnvironmentConfig } from '../common/config/environment';

export class SpService {
  protected context: WebPartContext;
  protected siteUrl: string;
  protected spHttpClient: SPHttpClient;
  protected httpClient: HttpClient;
  protected config: IEnvironmentConfig;

  constructor(context: WebPartContext) {
    this.context = context;
    this.siteUrl = context.pageContext.web.absoluteUrl;
    this.spHttpClient = context.spHttpClient;
    this.httpClient = context.httpClient;
    this.config = getConfig(this.siteUrl);
  }

  protected listUrl(listName: string): string {
    return `${this.siteUrl}/_api/web/lists/getByTitle('${encodeURIComponent(listName)}')`;
  }

  protected itemsUrl(listName: string, query: string = ''): string {
    return `${this.listUrl(listName)}/items${query}`;
  }

  protected get jsonHeaders(): Record<string, string> {
    return {
      'Accept': 'application/json;odata=nometadata',
      'Content-type': 'application/json;odata=nometadata',
      'odata-version': '',
    };
  }

  protected async getItems<T>(url: string): Promise<T[]> {
    const res: SPHttpClientResponse = await this.spHttpClient.get(url, SPHttpClient.configurations.v1);
    if (!res.ok) throw new Error(await parseSpError(res, res.status));
    const data = await res.json() as { value: T[] };
    return data.value || [];
  }

  protected async postItem<T>(url: string, body: object): Promise<T> {
    const res: SPHttpClientResponse = await this.spHttpClient.post(url, SPHttpClient.configurations.v1, {
      headers: this.jsonHeaders,
      body: JSON.stringify(body),
    });
    if (!res.ok) throw new Error(await parseSpError(res, res.status));
    return res.json() as Promise<T>;
  }

  protected async updateItem(listName: string, id: number, body: object): Promise<void> {
    const url = `${this.itemsUrl(listName)}(${id})`;
    const res: SPHttpClientResponse = await this.spHttpClient.fetch(url, SPHttpClient.configurations.v1, {
      method: 'PATCH',
      headers: {
        ...this.jsonHeaders,
        'IF-MATCH': '*',
        'X-HTTP-Method': 'MERGE',
      },
      body: JSON.stringify(body),
    });
    if (!res.ok && res.status !== 204) throw new Error(await parseSpError(res, res.status));
  }

  protected async deleteItem(listName: string, id: number): Promise<void> {
    const url = `${this.itemsUrl(listName)}(${id})`;
    const res: SPHttpClientResponse = await this.spHttpClient.fetch(url, SPHttpClient.configurations.v1, {
      method: 'DELETE',
      headers: { ...this.jsonHeaders, 'IF-MATCH': '*' },
    });
    if (!res.ok && res.status !== 204) throw new Error(await parseSpError(res, res.status));
  }

  // ── External API helpers ─────────────────────────────────────────────────

  protected apiUrl(path: string): string {
    return `${this.config.apiBaseUrl}/${path.replace(/^\//, '')}`;
  }

  protected async apiGet<T>(path: string): Promise<T> {
    const res: HttpClientResponse = await this.httpClient.get(
      this.apiUrl(path),
      HttpClient.configurations.v1,
      { headers: { 'Accept': 'application/json' } }
    );
    if (!res.ok) throw new Error(`API error ${res.status}: ${path}`);
    return res.json() as Promise<T>;
  }

  protected async apiPost<T>(path: string, body: object): Promise<T> {
    const res: HttpClientResponse = await this.httpClient.post(
      this.apiUrl(path),
      HttpClient.configurations.v1,
      {
        headers: { 'Accept': 'application/json', 'Content-type': 'application/json' },
        body: JSON.stringify(body),
      }
    );
    if (!res.ok) throw new Error(`API error ${res.status}: ${path}`);
    return res.json() as Promise<T>;
  }
}
