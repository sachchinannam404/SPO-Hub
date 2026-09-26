/**
 * List-Based Accordion Renderer
 */

import * as React from 'react';
import { useState, useMemo, useCallback } from 'react';
import {
  Accordion as FluentAccordion,
  AccordionItem,
  AccordionHeader,
  AccordionPanel
} from '@fluentui/react-accordion'; // fallback note: if not available use custom
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Stack } from '@fluentui/react/lib/Stack';
import { IconButton } from '@fluentui/react/lib/Button';
import { Text } from '@fluentui/react/lib/Text';
import { AccordionItem as AccordionDataItem } from '../../models/ListItem';
import { IDisplayConfig, IBehaviorConfig } from '../../models/Configuration';
import { LoadingState } from '../common/LoadingState/LoadingState';
import { EmptyState } from '../common/EmptyState/EmptyState';
import { ErrorState } from '../common/ErrorState/ErrorState';
import { UnauthorizedState } from '../common/UnauthorizedState/UnauthorizedState';
import { LoadState } from '../../hooks/useListData';
import styles from './Accordion.module.scss';

export interface IAccordionProps {
  items: AccordionDataItem[];
  state: LoadState;
  errorMessage?: string;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  onRetry?: () => void;
}

export const AccordionRenderer: React.FC<IAccordionProps> = ({
  items,
  state,
  errorMessage,
  display,
  behavior,
  onRetry
}) => {
  const [searchText, setSearchText] = useState('');
  const [expandedKeys, setExpandedKeys] = useState<string[]>(() => {
    if (behavior.expandFirstItem && items.length > 0) {
      return [String(items[0].id)];
    }
    return [];
  });

  const filtered = useMemo(() => {
    if (!searchText) return items;
    const q = searchText.toLowerCase();
    return items.filter(
      i =>
        (i.title || '').toLowerCase().indexOf(q) > -1 ||
        (i.content || '').toLowerCase().indexOf(q) > -1 ||
        (i.category || '').toLowerCase().indexOf(q) > -1
    );
  }, [items, searchText]);

  const expandAll = useCallback(() => {
    setExpandedKeys(filtered.map(i => String(i.id)));
  }, [filtered]);

  const collapseAll = useCallback(() => {
    setExpandedKeys([]);
  }, []);

  const toggle = useCallback(
    (key: string) => {
      setExpandedKeys(prev => {
        if (behavior.allowMultipleExpanded) {
          return prev.indexOf(key) > -1
            ? prev.filter(k => k !== key)
            : [...prev, key];
        }
        return prev.indexOf(key) > -1 ? [] : [key];
      });
    },
    [behavior.allowMultipleExpanded]
  );

  if (state === 'loading' || state === 'idle') {
    return <LoadingState label="Loading content..." />;
  }
  if (state === 'unauthorized') return <UnauthorizedState />;
  if (state === 'configError' || state === 'error') {
    return (
      <ErrorState
        message={errorMessage}
        onRetry={onRetry}
        isConfigError={state === 'configError'}
      />
    );
  }
  if (state === 'empty') {
    return <EmptyState message="No content is currently available." />;
  }

  return (
    <div className={styles.accordionRoot}>
      <Stack horizontal horizontalAlign="space-between" verticalAlign="center" styles={{ root: { marginBottom: 12 } }}>
        {behavior.enableSearch && (
          <SearchBox
            placeholder="Search..."
            onChange={(_, v) => setSearchText(v || '')}
            value={searchText}
            styles={{ root: { minWidth: 220 } }}
            ariaLabel="Search accordion items"
          />
        )}
        <Stack horizontal tokens={{ childrenGap: 4 }}>
          <IconButton
            iconProps={{ iconName: 'ExploreContent' }}
            title="Expand all"
            ariaLabel="Expand all"
            onClick={expandAll}
          />
          <IconButton
            iconProps={{ iconName: 'CollapseContent' }}
            title="Collapse all"
            ariaLabel="Collapse all"
            onClick={collapseAll}
          />
        </Stack>
      </Stack>

      {filtered.length === 0 ? (
        <EmptyState message="No items match your search." iconName="Search" />
      ) : (
        <div className={styles.accordionList} role="region" aria-label={display.title || 'Accordion'}>
          {filtered.map(item => {
            const key = String(item.id);
            const isExpanded = expandedKeys.indexOf(key) > -1;
            return (
              <div key={key} className={styles.accordionItem}>
                <button
                  type="button"
                  className={styles.accordionHeader}
                  aria-expanded={isExpanded}
                  aria-controls={`panel-${key}`}
                  id={`header-${key}`}
                  onClick={() => toggle(key)}
                >
                  <Text className={styles.headerText}>{item.title}</Text>
                  <span className={styles.chevron} aria-hidden="true">
                    {isExpanded ? '▾' : '▸'}
                  </span>
                </button>
                {isExpanded && (
                  <div
                    id={`panel-${key}`}
                    role="region"
                    aria-labelledby={`header-${key}`}
                    className={styles.accordionPanel}
                  >
                    {item.content ? (
                      <div
                        className={styles.richContent}
                        dangerouslySetInnerHTML={{ __html: sanitizeHtml(item.content) }}
                      />
                    ) : (
                      <Text>{item.description}</Text>
                    )}
                    {item.link && (
                      <a href={item.link} target="_blank" rel="noopener noreferrer">
                        Learn more
                      </a>
                    )}
                  </div>
                )}
              </div>
            );
          })}
        </div>
      )}
    </div>
  );
};

/** Minimal sanitizer – production should use a vetted library (DOMPurify) */
function sanitizeHtml(html: string): string {
  if (!html) return '';
  return html
    .replace(/<script\b[^<]*(?:(?!<\/script>)<[^<]*)*<\/script>/gi, '')
    .replace(/on\w+="[^"]*"/gi, '')
    .replace(/on\w+='[^']*'/gi, '');
}

export default AccordionRenderer;
