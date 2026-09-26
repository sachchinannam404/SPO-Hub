/**
 * List-Based Tiles Renderer
 * Responsive grid of configurable tiles driven by SharePoint list data.
 */

import * as React from 'react';
import { useMemo, useState, useCallback } from 'react';
import { Stack } from '@fluentui/react/lib/Stack';
import { SearchBox } from '@fluentui/react/lib/SearchBox';
import { Dropdown, IDropdownOption } from '@fluentui/react/lib/Dropdown';
import { FocusZone } from '@fluentui/react/lib/FocusZone';
import { TileItem } from '../../models/ListItem';
import { IDisplayConfig, IBehaviorConfig } from '../../models/Configuration';
import { LoadingState } from '../common/LoadingState/LoadingState';
import { EmptyState } from '../common/EmptyState/EmptyState';
import { ErrorState } from '../common/ErrorState/ErrorState';
import { UnauthorizedState } from '../common/UnauthorizedState/UnauthorizedState';
import { TileCard } from './TileCard';
import { LoadState } from '../../hooks/useListData';
import styles from './Tiles.module.scss';

export interface ITilesProps {
  items: TileItem[];
  state: LoadState;
  errorMessage?: string;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  onRetry?: () => void;
  onItemClick?: (item: TileItem) => void;
  webPartWidth?: number;
}

export const Tiles: React.FC<ITilesProps> = ({
  items,
  state,
  errorMessage,
  display,
  behavior,
  onRetry,
  onItemClick,
  webPartWidth = 1200
}) => {
  const [searchText, setSearchText] = useState('');
  const [selectedCategory, setSelectedCategory] = useState<string>('all');

  const categories = useMemo(() => {
    const set = new Set<string>();
    items.forEach(i => {
      if (i.category) set.add(i.category);
    });
    return Array.from(set).sort();
  }, [items]);

  const categoryOptions: IDropdownOption[] = useMemo(() => {
    return [
      { key: 'all', text: 'All categories' },
      ...categories.map(c => ({ key: c, text: c }))
    ];
  }, [categories]);

  const filtered = useMemo(() => {
    let result = items;
    if (searchText) {
      const q = searchText.toLowerCase();
      result = result.filter(
        i =>
          (i.title || '').toLowerCase().indexOf(q) > -1 ||
          (i.description || '').toLowerCase().indexOf(q) > -1 ||
          (i.category || '').toLowerCase().indexOf(q) > -1
      );
    }
    if (selectedCategory && selectedCategory !== 'all') {
      result = result.filter(i => i.category === selectedCategory);
    }
    return result;
  }, [items, searchText, selectedCategory]);

  const columns = useMemo(() => {
    // Responsive columns based on available width
    if (webPartWidth < 480) return 1;
    if (webPartWidth < 768) return 2;
    if (webPartWidth < 1024) return Math.min(display.columns || 3, 3);
    return display.columns || 4;
  }, [webPartWidth, display.columns]);

  const handleSearch = useCallback((_: any, newValue?: string) => {
    setSearchText(newValue || '');
  }, []);

  if (state === 'loading' || state === 'idle') {
    return <LoadingState label="Loading tiles..." />;
  }
  if (state === 'unauthorized') {
    return <UnauthorizedState />;
  }
  if (state === 'configError' || state === 'error') {
    return (
      <ErrorState
        message={errorMessage}
        onRetry={onRetry}
        isConfigError={state === 'configError'}
      />
    );
  }
  if (state === 'empty' || (filtered.length === 0 && items.length === 0)) {
    return <EmptyState message="No tiles are currently available." />;
  }

  return (
    <div className={styles.tilesRoot}>
      {(behavior.enableSearch || behavior.enableFilter) && (
        <Stack
          horizontal
          wrap
          tokens={{ childrenGap: 12 }}
          styles={{ root: { marginBottom: 16 } }}
        >
          {behavior.enableSearch && (
            <SearchBox
              placeholder="Search..."
              onChange={handleSearch}
              value={searchText}
              styles={{ root: { minWidth: 200, maxWidth: 320 } }}
              ariaLabel="Search tiles"
            />
          )}
          {behavior.enableFilter && categories.length > 0 && (
            <Dropdown
              options={categoryOptions}
              selectedKey={selectedCategory}
              onChange={(_, opt) => setSelectedCategory((opt?.key as string) || 'all')}
              styles={{ root: { minWidth: 180 } }}
              ariaLabel="Filter by category"
            />
          )}
        </Stack>
      )}

      {filtered.length === 0 ? (
        <EmptyState message="No tiles match your search or filter." iconName="Search" />
      ) : (
        <FocusZone>
          <div
            className={styles.tilesGrid}
            style={{
              gridTemplateColumns: `repeat(${columns}, minmax(0, 1fr))`
            }}
            role="list"
          >
            {filtered.map(item => (
              <div key={item.id} role="listitem">
                <TileCard
                  item={item}
                  display={display}
                  openInNewTab={behavior.openLinksInNewTab}
                  onClick={onItemClick}
                />
              </div>
            ))}
          </div>
        </FocusZone>
      )}
    </div>
  );
};

export default Tiles;
