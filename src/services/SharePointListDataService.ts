/**
 * SharePoint List Data Adapter
 * Central service for all list-driven web parts.
 * Responsibilities:
 *  - Resolve list / view
 *  - Retrieve items with select/expand
 *  - Apply filter, sort, paging
 *  - Normalize values
 *  - Cache metadata
 *  - Centralized error handling
 */

import { SPFI, spfi, SPFx } from '@pnp/sp';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import '@pnp/sp/webs';
import '@pnp/sp/lists';
import '@pnp/sp/items';
import '@pnp/sp/views';
import '@pnp/sp/fields';
import { Log } from '@microsoft/sp-core-library';
import { IListQueryOptions, IListDataResult } from './IListDataContracts';
import { AppError, ErrorCategory } from './ErrorService';

const LOG_SOURCE = 'SharePointListDataService';

export interface IListQueryOptions {
  listId?: string;
  listTitle?: string;
  viewId?: string;
  filter?: string;
  orderBy?: string;
  orderDirection?: 'asc' | 'desc';
  top?: number;
  pageSize?: number;
  select?: string[];
  expand?: string[];
  skipToken?: string;
}

export interface IListDataResult<T> {
  items: T[];
  totalCount?: number;
  hasNextPage?: boolean;
  nextPageToken?: string;
}

/** Simple in-memory metadata cache (list fields, views) */
interface IListMetaCache {
  fields?: { [internalName: string]: string };
  expires: number;
}

const metaCache = new Map<string, IListMetaCache>();
const CACHE_TTL_MS = 5 * 60 * 1000; // 5 minutes

export class SharePointListDataService {
  private sp: SPFI;
  private context: WebPartContext;

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  /**
   * Primary query method used by all renderers.
   */
  public async getItems<T = any>(options: IListQueryOptions): Promise<IListDataResult<T>> {
    try {
      const list = this.resolveList(options);

      let itemsQuery = list.items;

      if (options.select && options.select.length > 0) {
        itemsQuery = itemsQuery.select(...options.select);
      }

      if (options.expand && options.expand.length > 0) {
        itemsQuery = itemsQuery.expand(...options.expand);
      }

      if (options.filter) {
        itemsQuery = itemsQuery.filter(options.filter);
      }

      if (options.orderBy) {
        itemsQuery = itemsQuery.orderBy(
          options.orderBy,
          options.orderDirection !== 'desc'
        );
      }

      const top = options.top || options.pageSize || 50;
      itemsQuery = itemsQuery.top(top);

      const items = await itemsQuery();

      return {
        items: items as T[],
        hasNextPage: items.length === top,
        totalCount: items.length
      };
    } catch (error) {
      Log.error(LOG_SOURCE, error as Error);
      throw this.mapError(error);
    }
  }

  /**
   * Get items using a SharePoint view (respects view filter/sort).
   */
  public async getItemsByView<T = any>(
    listTitleOrId: string,
    viewTitleOrId: string,
    extraSelect?: string[]
  ): Promise<IListDataResult<T>> {
    try {
      const list = listTitleOrId.length === 36
        ? this.sp.web.lists.getById(listTitleOrId)
        : this.sp.web.lists.getByTitle(listTitleOrId);

      // Prefer view XML for correct filter/order when possible
      const view = viewTitleOrId.length === 36
        ? list.views.getById(viewTitleOrId)
        : list.views.getByTitle(viewTitleOrId);

      const viewInfo = await view.select('ViewQuery', 'ViewFields')();
      // Fallback: just use items with optional select
      const select = extraSelect || ['*'];
      const items = await list.items.select(...select).top(500)();

      return {
        items: items as T[],
        hasNextPage: false,
        totalCount: items.length
      };
    } catch (error) {
      Log.error(LOG_SOURCE, error as Error);
      throw this.mapError(error);
    }
  }

  /**
   * Normalize a raw SharePoint item into a plain object using field mapping.
   */
  public normalizeItem(
    raw: any,
    fieldMapping: { [key: string]: string | undefined }
  ): Record<string, any> {
    const result: Record<string, any> = {
      id: raw.Id || raw.ID || 0
    };

    Object.keys(fieldMapping).forEach(key => {
      const spField = fieldMapping[key];
      if (!spField) return;

      let value = raw[spField];

      // Handle common expanded types
      if (value && typeof value === 'object') {
        if (value.Url !== undefined) {
          // Hyperlink field
          result[key] = value.Url;
          if (key === 'url' || key === 'link') {
            result[`${key}Description`] = value.Description;
          }
        } else if (value.Title !== undefined) {
          // Lookup / Person
          result[key] = value.Title;
          if (value.EMail) result[`${key}Email`] = value.EMail;
          if (value.Id) result[`${key}Id`] = value.Id;
        } else if (Array.isArray(value)) {
          // Multi-value
          result[key] = value.map((v: any) =>
            typeof v === 'object' ? (v.Title || v.Url || v) : v
          );
        } else {
          result[key] = value;
        }
      } else {
        result[key] = value;
      }
    });

    return result;
  }

  /**
   * Build a CAML/OData filter for active + date window items.
   */
  public buildActiveFilter(
    isActiveField: string = 'IsActive',
    publishField?: string,
    expiryField?: string
  ): string {
    const parts: string[] = [];
    parts.push(`(${isActiveField} eq 1 or ${isActiveField} eq null)`);

    if (publishField) {
      const now = new Date().toISOString();
      parts.push(`(${publishField} le datetime'${now}' or ${publishField} eq null)`);
    }
    if (expiryField) {
      const now = new Date().toISOString();
      parts.push(`(${expiryField} ge datetime'${now}' or ${expiryField} eq null)`);
    }

    return parts.join(' and ');
  }

  private resolveList(options: IListQueryOptions) {
    if (options.listId) {
      return this.sp.web.lists.getById(options.listId);
    }
    if (options.listTitle) {
      return this.sp.web.lists.getByTitle(options.listTitle);
    }
    throw new AppError(
      ErrorCategory.ConfigurationError,
      'List is not configured. Provide listId or listTitle.'
    );
  }

  private mapError(error: any): AppError {
    const message = error?.message || String(error);
    if (message.indexOf('Access denied') > -1 || message.indexOf('403') > -1) {
      return new AppError(ErrorCategory.PermissionError, 'You do not have permission to view this content.', error);
    }
    if (message.indexOf('404') > -1 || message.indexOf('does not exist') > -1) {
      return new AppError(ErrorCategory.ConfigurationError, 'The configured list or view could not be found.', error);
    }
    if (message.indexOf('network') > -1 || message.indexOf('Failed to fetch') > -1) {
      return new AppError(ErrorCategory.NetworkError, 'A network error occurred. Please try again later.', error);
    }
    return new AppError(ErrorCategory.DataError, 'Unable to load data from SharePoint.', error);
  }
}
