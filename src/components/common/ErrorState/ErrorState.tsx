import * as React from 'react';
import { MessageBar, MessageBarType } from '@fluentui/react/lib/MessageBar';
import { PrimaryButton } from '@fluentui/react/lib/Button';
import { Stack } from '@fluentui/react/lib/Stack';

export interface IErrorStateProps {
  message?: string;
  onRetry?: () => void;
  isConfigError?: boolean;
}

export const ErrorState: React.FC<IErrorStateProps> = ({
  message = 'Something went wrong while loading this content.',
  onRetry,
  isConfigError = false
}) => {
  return (
    <Stack tokens={{ childrenGap: 12 }} styles={{ root: { padding: 16 } }}>
      <MessageBar
        messageBarType={isConfigError ? MessageBarType.warning : MessageBarType.error}
        isMultiline
        role="alert"
      >
        {message}
      </MessageBar>
      {onRetry && !isConfigError && (
        <PrimaryButton text="Try again" onClick={onRetry} />
      )}
    </Stack>
  );
};

export default ErrorState;
