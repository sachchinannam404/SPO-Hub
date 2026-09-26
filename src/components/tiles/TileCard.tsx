import * as React from 'react';
import { useCallback } from 'react';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';
import { TileItem } from '../../models/ListItem';
import { IDisplayConfig } from '../../models/Configuration';
import styles from './Tiles.module.scss';

export interface ITileCardProps {
  item: TileItem;
  display: IDisplayConfig;
  openInNewTab?: boolean;
  onClick?: (item: TileItem) => void;
}

export const TileCard: React.FC<ITileCardProps> = ({
  item,
  display,
  openInNewTab = true,
  onClick
}) => {
  const handleClick = useCallback(
    (e: React.MouseEvent | React.KeyboardEvent) => {
      if (onClick) {
        onClick(item);
      }
      if (item.url) {
        // Navigation is handled by the anchor; telemetry via onClick
      } else {
        e.preventDefault();
      }
    },
    [item, onClick]
  );

  const handleKeyDown = useCallback(
    (e: React.KeyboardEvent) => {
      if (e.key === 'Enter' || e.key === ' ') {
        handleClick(e);
      }
    },
    [handleClick]
  );

  const content = (
    <>
      {(item.imageUrl || item.iconName) && (
        <div className={styles.tileImage} aria-hidden={!!item.iconName}>
          {item.imageUrl ? (
            <img
              src={item.imageUrl}
              alt={item.altText || item.title || ''}
              loading="lazy"
            />
          ) : (
            <Icon
              iconName={item.iconName || 'Link'}
              styles={{
                root: {
                  fontSize: 36,
                  color: item.iconColor || '#0078d4'
                }
              }}
            />
          )}
        </div>
      )}
      <div className={styles.tileBody}>
        <Text className={styles.tileTitle} as="h3">
          {item.title}
        </Text>
        {display.showDescription && item.description && (
          <Text className={styles.tileDescription}>{item.description}</Text>
        )}
        {display.showCategory && item.category && (
          <Text className={styles.tileCategory} variant="small">
            {item.category}
          </Text>
        )}
      </div>
    </>
  );

  const className = [
    styles.tileCard,
    display.cardSize === 'small' ? styles.small : '',
    display.cardSize === 'large' ? styles.large : ''
  ]
    .filter(Boolean)
    .join(' ');

  if (item.url) {
    return (
      <a
        className={className}
        href={item.url}
        target={openInNewTab ? '_blank' : undefined}
        rel={openInNewTab ? 'noopener noreferrer' : undefined}
        onClick={handleClick}
        onKeyDown={handleKeyDown}
        aria-label={item.title}
      >
        {content}
      </a>
    );
  }

  return (
    <div
      className={className}
      role="button"
      tabIndex={0}
      onClick={handleClick}
      onKeyDown={handleKeyDown}
      aria-label={item.title}
    >
      {content}
    </div>
  );
};

export default TileCard;
