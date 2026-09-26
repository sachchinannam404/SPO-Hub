import * as React from 'react';
import { Spinner, SpinnerSize } from '@fluentui/react/lib/Spinner';
import { Stack } from '@fluentui/react/lib/Stack';

export interface ILoadingStateProps {
  label?: string;
}

export const LoadingState: React.FC<ILoadingStateProps> = ({ label = 'Loading...' }) => {
  return (
    <Stack horizontalAlign="center" verticalAlign="center" styles={{ root: { minHeight: 120, padding: 24 } }}>
      <Spinner size={SpinnerSize.medium} label={label} ariaLive="polite" />
    </Stack>
  );
};

export default LoadingState;
