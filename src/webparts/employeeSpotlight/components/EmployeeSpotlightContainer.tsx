import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IDataSourceConfig,
  IDisplayConfig,
  IBehaviorConfig,
  IPersonalizationConfig
} from '../../../models/Configuration';
import { EmployeeSpotlightItem, IFieldMapping } from '../../../models/ListItem';
import { useListData } from '../../../hooks/useListData';
import { EmployeeSpotlight } from '../../../components/employeeSpotlight/EmployeeSpotlight';

export interface IEmployeeSpotlightContainerProps {
  context: WebPartContext;
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  personalization: IPersonalizationConfig;
  webPartWidth: number;
}

function mapEmployee(raw: any, _mapping: IFieldMapping): EmployeeSpotlightItem {
  return {
    id: raw.id || 0,
    title: raw.title || raw.employeeName || '',
    employeeName: raw.employeeName || raw.title || '',
    jobTitle: raw.jobTitle,
    department: raw.department,
    location: raw.location,
    imageUrl: raw.imageUrl || raw.image || raw.photo,
    biography: raw.biography,
    achievement: raw.achievement,
    recognitionDate: raw.recognitionDate,
    profileUrl: raw.profileUrl,
    isFeatured: !!raw.isFeatured,
    category: raw.category,
    isActive: raw.isActive !== false,
    audience: raw.audience,
    employeeEmail: raw.employeeEmail
  };
}

export const EmployeeSpotlightContainer: React.FC<IEmployeeSpotlightContainerProps> = (props) => {
  const { context, dataSource, display, behavior, personalization, webPartWidth } = props;

  const { items, state, errorMessage, refresh } = useListData<EmployeeSpotlightItem>(
    context,
    dataSource,
    'EmployeeSpotlight',
    mapEmployee,
    { enableAudience: personalization.enableAudienceTargeting }
  );

  return (
    <EmployeeSpotlight
      items={items}
      state={state}
      errorMessage={errorMessage}
      display={display}
      behavior={behavior}
      onRetry={refresh}
      webPartWidth={webPartWidth}
    />
  );
};

export default EmployeeSpotlightContainer;
