import * as React from 'react';
import { WebPartContext } from '@microsoft/sp-webpart-base';
import {
  IDataSourceConfig,
  IDisplayConfig,
  IBehaviorConfig,
  IPersonalizationConfig
} from '../../../models/Configuration';
import { QuestionItem, IFieldMapping, QuestionStatus } from '../../../models/ListItem';
import { useListData } from '../../../hooks/useListData';
import { QuestionsAnswers } from '../../../components/questionsAnswers/QuestionsAnswers';

export interface IQuestionsAnswersContainerProps {
  context: WebPartContext;
  dataSource: IDataSourceConfig;
  display: IDisplayConfig;
  behavior: IBehaviorConfig;
  personalization: IPersonalizationConfig;
}

function mapQuestion(raw: any, _mapping: IFieldMapping): QuestionItem {
  let tags: string[] | undefined;
  if (Array.isArray(raw.tags)) {
    tags = raw.tags;
  } else if (typeof raw.tags === 'string' && raw.tags) {
    tags = raw.tags.split(/[,;]/).map((t: string) => t.trim()).filter(Boolean);
  }

  return {
    id: raw.id || 0,
    title: raw.title || raw.question || '',
    question: raw.question || raw.title || '',
    answer: raw.answer,
    category: raw.category,
    tags,
    askedBy: raw.askedBy,
    answeredBy: raw.answeredBy,
    questionDate: raw.questionDate,
    answerDate: raw.answerDate,
    status: (raw.status as QuestionStatus) || undefined,
    isFeatured: !!raw.isFeatured,
    isActive: raw.isActive !== false,
    audience: raw.audience
  };
}

export const QuestionsAnswersContainer: React.FC<IQuestionsAnswersContainerProps> = (props) => {
  const { context, dataSource, display, behavior, personalization } = props;

  const { items, state, errorMessage, refresh } = useListData<QuestionItem>(
    context,
    dataSource,
    'QuestionsAnswers',
    mapQuestion,
    { enableAudience: personalization.enableAudienceTargeting }
  );

  return (
    <QuestionsAnswers
      items={items}
      state={state}
      errorMessage={errorMessage}
      display={display}
      behavior={behavior}
      onRetry={refresh}
    />
  );
};

export default QuestionsAnswersContainer;
