import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Text } from '@fluentui/react/lib/Text';
import { Icon } from '@fluentui/react/lib/Icon';

export interface IEmptyStateProps {
  message?: string;
  iconName?: string;
}

export const EmptyState: React.FC<IEmptyStateProps> = ({
  message = 'No content is currently available.',
  iconName = 'Info'
}) => {
  return (
    <Stack
      horizontalAlign="center"
      verticalAlign="center"
      tokens={{ childrenGap: 12 }}
      styles={{ root: { minHeight: 120, padding: 24 } }}
      role="status"
      aria-live="polite"
    >
      <Icon iconName={iconName} styles={{ root: { fontSize: 32, color: '#605e5c' } }} aria-hidden="true" />
      <Text variant="medium">{message}</Text>
    </Stack>
  );
};

export default EmptyState;
