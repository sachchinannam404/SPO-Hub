import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { Stack } from '@fluentui/react/lib/Stack';
import { Icon } from '@fluentui/react/lib/Icon';
import { Text } from '@fluentui/react/lib/Text';

export interface IUnauthorizedStateProps {
  message?: string;
}

export const UnauthorizedState: React.FC<IUnauthorizedStateProps> = ({
  message = "You don't have permission to view this content."
}) => {
  return (
    <Stack
      horizontalAlign="center"
      verticalAlign="center"
      tokens={{ childrenGap: 12 }}
      styles={{ root: { minHeight: 120, padding: 24 } }}
      role="alert"
    >
      <Icon iconName="Lock" styles={{ root: { fontSize: 32, color: '#a4262c' } }} aria-hidden="true" />
      <Text variant="medium">{message}</Text>
    </Stack>
  );
};

export default UnauthorizedState;
