/**
 * List-Based Vertical Timeline Renderer
 * Built as a dedicated component using the common List Data Adapter.
 */

import * as React from 'react';
import { useMemo } from 'react';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';
import { Link } from '@fluentui/react/lib/Link';
import { TimelineItem, TimelineStatus } from '../../models/ListItem';
import { IDisplayConfig, IBehaviorConfig } from '../../models/Configuration';
import { LoadingState } from '../common/LoadingState/LoadingState';
import { EmptyState } from '../common/EmptyState/EmptyState';
import { ErrorState } from '../common/ErrorState/ErrorState';
import { UnauthorizedState } from '../common/UnauthorizedState/UnauthorizedState';
import { LoadState } from '../../hooks/useListData';
import styles from './Timeline.module.scss';

export interface ITimelineProps {
  items: TimelineItem[];
  state: LoadState;
  errorMessage?: string;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  onRetry?: () => void;
  compact?: boolean;
}

const STATUS_ICON: Record<TimelineStatus, string> = {
  Completed: 'CompletedSolid',
  'In Progress': 'ProgressRingDots',
  Upcoming: 'Calendar',
  'On Hold': 'Pause',
  Cancelled: 'StatusErrorFull'
};

const STATUS_CLASS: Record<TimelineStatus, string> = {
  Completed: styles.statusCompleted,
  'In Progress': styles.statusInProgress,
  Upcoming: styles.statusUpcoming,
  'On Hold': styles.statusOnHold,
  Cancelled: styles.statusCancelled
};

export const Timeline: React.FC<ITimelineProps> = ({
  items,
  state,
  errorMessage,
  display,
  onRetry,
  compact = false
}) => {
  const sorted = useMemo(() => {
    return [...items].sort((a, b) => {
      if (a.sequence != null && b.sequence != null) {
        return a.sequence - b.sequence;
      }
      const da = a.startDate ? new Date(a.startDate).getTime() : 0;
      const db = b.startDate ? new Date(b.startDate).getTime() : 0;
      return da - db;
    });
  }, [items]);

  if (state === 'loading' || state === 'idle') {
    return <LoadingState label="Loading timeline..." />;
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
    return <EmptyState message="No timeline milestones are currently available." />;
  }

  return (
    <div className={styles.timelineRoot}>
      {display.title && (
        <Text as="h2" variant="xLarge" styles={{ root: { marginBottom: 20, fontWeight: 600 } }}>
          {display.title}
        </Text>
      )}
      <ol className={`${styles.timeline} ${compact ? styles.compact : ''}`} aria-label="Timeline">
        {sorted.map((item, index) => {
          const status = (item.status as TimelineStatus) || 'Upcoming';
          const iconName = item.icon || STATUS_ICON[status] || 'CircleShapeSolid';
          return (
            <li key={item.id} className={styles.timelineItem}>
              <div className={`${styles.marker} ${STATUS_CLASS[status] || ''}`}>
                <Icon iconName={iconName} aria-hidden="true" />
              </div>
              {index < sorted.length - 1 && <div className={styles.line} aria-hidden="true" />}
              <div className={styles.content}>
                <div className={styles.meta}>
                  {item.startDate && (
                    <time dateTime={item.startDate} className={styles.date}>
                      {formatDate(item.startDate)}
                      {item.endDate ? ` – ${formatDate(item.endDate)}` : ''}
                    </time>
                  )}
                  {item.status && (
                    <span className={`${styles.badge} ${STATUS_CLASS[status] || ''}`}>
                      {item.status}
                    </span>
                  )}
                  {item.category && (
                    <span className={styles.category}>{item.category}</span>
                  )}
                </div>
                <Text as="h3" className={styles.itemTitle}>
                  {item.title}
                </Text>
                {item.description && (
                  <Text className={styles.description}>{item.description}</Text>
                )}
                <div className={styles.footer}>
                  {item.owner && (
                    <Text variant="small" className={styles.owner}>
                      Owner: {item.owner}
                    </Text>
                  )}
                  {item.link && (
                    <Link href={item.link} target="_blank" rel="noopener noreferrer">
                      Details
                    </Link>
                  )}
                </div>
              </div>
            </li>
          );
        })}
      </ol>
    </div>
  );
};

function formatDate(iso: string): string {
  try {
    return new Date(iso).toLocaleDateString(undefined, {
      year: 'numeric',
      month: 'short',
      day: 'numeric'
    });
  } catch {
    return iso;
  }
}

export default Timeline;
