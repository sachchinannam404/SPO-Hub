/**
 * Normalized Component Model
 * Common base types for all list-driven renderers.
 */

export interface BaseListItem {
  id: number;
  title: string;
  isActive?: boolean;
  displayOrder?: number;
  category?: string;
  audience?: string[];
  publishDate?: string;
  expiryDate?: string;
}

export interface LinkItem extends BaseListItem {
  url?: string;
  openInNewTab?: boolean;
  description?: string;
}

export interface ImageItem extends BaseListItem {
  imageUrl?: string;
  altText?: string;
}

export interface TileItem extends LinkItem, ImageItem {
  iconName?: string;
  iconColor?: string;
  backgroundColor?: string;
  target?: string;
}

export interface AccordionItem extends BaseListItem {
  description?: string;
  content?: string;
  link?: string;
  lastUpdated?: string;
}

export interface QuoteItem extends BaseListItem, ImageItem {
  quote: string;
  author?: string;
  role?: string;
  department?: string;
}

export interface EmployeeSpotlightItem extends BaseListItem, ImageItem {
  employeeName: string;
  jobTitle?: string;
  department?: string;
  location?: string;
  biography?: string;
  achievement?: string;
  recognitionDate?: string;
  profileUrl?: string;
  isFeatured?: boolean;
  employeeEmail?: string;
}

export interface QuestionItem extends BaseListItem {
  question: string;
  answer?: string;
  tags?: string[];
  askedBy?: string;
  answeredBy?: string;
  questionDate?: string;
  answerDate?: string;
  status?: QuestionStatus;
  isFeatured?: boolean;
}

export type QuestionStatus =
  | 'New'
  | 'In Review'
  | 'Answered'
  | 'Closed'
  | 'Archived';

export interface TimelineItem extends BaseListItem {
  description?: string;
  startDate?: string;
  endDate?: string;
  status?: TimelineStatus;
  icon?: string;
  imageUrl?: string;
  link?: string;
  owner?: string;
  sequence?: number;
}

export type TimelineStatus =
  | 'Completed'
  | 'In Progress'
  | 'Upcoming'
  | 'On Hold'
  | 'Cancelled';

export interface ChartDataPoint {
  category: string;
  value: number;
  series?: string;
  color?: string;
}

export interface ChartSeriesConfig {
  field: string;
  aggregation: 'sum' | 'count' | 'avg' | 'min' | 'max';
  label?: string;
  color?: string;
}

export type ChartType =
  | 'bar'
  | 'horizontalBar'
  | 'line'
  | 'area'
  | 'pie'
  | 'doughnut'
  | 'scatter';

/** Field mapping configuration used by all list-driven web parts */
export interface IFieldMapping {
  titleField?: string;
  descriptionField?: string;
  contentField?: string;
  imageField?: string;
  linkField?: string;
  categoryField?: string;
  orderField?: string;
  isActiveField?: string;
  audienceField?: string;
  publishDateField?: string;
  expiryDateField?: string;
  // Quote specific
  quoteField?: string;
  authorField?: string;
  roleField?: string;
  // Employee specific
  employeeNameField?: string;
  jobTitleField?: string;
  departmentField?: string;
  locationField?: string;
  biographyField?: string;
  achievementField?: string;
  profileUrlField?: string;
  // Q&A specific
  questionField?: string;
  answerField?: string;
  tagsField?: string;
  statusField?: string;
  // Timeline specific
  startDateField?: string;
  endDateField?: string;
  statusTimelineField?: string;
  iconField?: string;
  ownerField?: string;
  sequenceField?: string;
}
