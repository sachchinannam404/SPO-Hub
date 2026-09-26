/**
 * React hook for list-driven data loading.
 * Encapsulates loading / error / empty states and refresh logic.
 */

import { useState, useEffect, useCallback, useRef } from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SharePointListDataService, IListQueryOptions } from '../services/SharePointListDataService';
import { AudienceService } from '../services/AudienceService';
import { TelemetryService } from '../services/TelemetryService';
import { AppError, getUserFacingMessage } from '../services/ErrorService';
import { BaseListItem, IFieldMapping } from '../models/ListItem';
import { IDataSourceConfig } from '../models/Configuration';

export type LoadState = 'idle' | 'loading' | 'success' | 'empty' | 'error' | 'unauthorized' | 'configError';

export interface UseListDataResult<T> {
  items: T[];
  state: LoadState;
  errorMessage?: string;
  refresh: () => void;
  isLoading: boolean;
}

export function useListData<T extends BaseListItem>(
  context: WebPartContext,
  dataSource: IDataSourceConfig,
  componentName: string,
  mapItem: (raw: any, mapping: IFieldMapping) => T,
  options?: {
    enableAudience?: boolean;
    extraSelect?: string[];
    extraExpand?: string[];
  }
): UseListDataResult<T> {
  const [items, setItems] = useState<T[]>([]);
  const [state, setState] = useState<LoadState>('idle');
  const [errorMessage, setErrorMessage] = useState<string | undefined>();
  const mounted = useRef(true);

  const dataService = useRef(new SharePointListDataService(context));
  const audienceService = useRef(new AudienceService(context));
  const telemetry = useRef(new TelemetryService(context, componentName));

  const load = useCallback(async () => {
    if (!dataSource?.listTitle && !dataSource?.listId) {
      setState('configError');
      setErrorMessage('This web part has not been configured correctly. Please review the data source configuration.');
      return;
    }

    setState('loading');
    setErrorMessage(undefined);

    try {
      const mapping = dataSource.fieldMapping || {};
      const selectFields = buildSelect(mapping, options?.extraSelect);
      const expandFields = buildExpand(mapping, options?.extraExpand);

      const query: IListQueryOptions = {
        listId: dataSource.listId,
        listTitle: dataSource.listTitle,
        viewId: dataSource.viewId,
        filter: dataSource.filter,
        orderBy: dataSource.orderBy || mapping.orderField || 'Title',
        orderDirection: dataSource.orderDirection || 'asc',
        top: dataSource.itemLimit || 50,
        select: selectFields,
        expand: expandFields
      };

      // Prefer active + date window when fields exist
      if (mapping.isActiveField && !query.filter) {
        query.filter = dataService.current.buildActiveFilter(
          mapping.isActiveField,
          mapping.publishDateField,
          mapping.expiryDateField
        );
      }

      const result = await dataService.current.getItems<any>(query);
      let normalized = result.items.map(raw =>
        mapItem(dataService.current.normalizeItem(raw, mapping as any), mapping)
      );

      // Audience targeting (client-side after reasonable server top)
      if (options?.enableAudience) {
        const userCtx = await audienceService.current.getCurrentUserContext();
        normalized = audienceService.current.filterByAudience(normalized, userCtx);
      }

      if (!mounted.current) return;

      if (normalized.length === 0) {
        setItems([]);
        setState('empty');
      } else {
        setItems(normalized);
        setState('success');
      }

      telemetry.current.trackWebPartLoaded({ itemCount: normalized.length });
    } catch (err) {
      if (!mounted.current) return;
      const message = getUserFacingMessage(err);
      setErrorMessage(message);
      if (err instanceof AppError) {
        if (err.category === 'PermissionError') setState('unauthorized');
        else if (err.category === 'ConfigurationError') setState('configError');
        else setState('error');
      } else {
        setState('error');
      }
    }
  }, [context, dataSource, componentName, mapItem, options]);

  useEffect(() => {
    mounted.current = true;
    load();
    return () => {
      mounted.current = false;
    };
  }, [load]);

  // Auto-refresh
  useEffect(() => {
    const seconds = dataSource.refreshIntervalSeconds || 0;
    if (seconds <= 0) return;
    const id = setInterval(load, seconds * 1000);
    return () => clearInterval(id);
  }, [dataSource.refreshIntervalSeconds, load]);

  return {
    items,
    state,
    errorMessage,
    refresh: load,
    isLoading: state === 'loading'
  };
}

function buildSelect(mapping: IFieldMapping, extra?: string[]): string[] {
  const fields = new Set<string>(['Id', 'Title']);
  Object.values(mapping).forEach(v => {
    if (v) fields.add(v);
  });
  (extra || []).forEach(f => fields.add(f));
  return Array.from(fields);
}

function buildExpand(mapping: IFieldMapping, extra?: string[]): string[] {
  const expand = new Set<string>();
  // Common person/lookup fields that need expand
  const candidates = [
    mapping.imageField,
    mapping.linkField,
    mapping.audienceField,
    mapping.ownerField
  ];
  candidates.forEach(f => {
    if (f) expand.add(f);
  });
  (extra || []).forEach(f => expand.add(f));
  return Array.from(expand);
}
