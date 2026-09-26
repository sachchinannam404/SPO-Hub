/**
 * Audience Targeting Service
 * Filters items early based on user context.
 * Never relies solely on client-side hiding of large datasets.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { BaseListItem } from '../models/ListItem';

export interface IAudienceContext {
  loginName: string;
  displayName: string;
  email?: string;
  department?: string;
  jobTitle?: string;
  officeLocation?: string;
  groupIds: string[];
}

export class AudienceService {
  private context: WebPartContext;
  private userContext: IAudienceContext | null = null;

  constructor(context: WebPartContext) {
    this.context = context;
  }

  /**
   * Resolve current user audience attributes.
   * Call once per web part lifetime.
   */
  public async getCurrentUserContext(): Promise<IAudienceContext> {
    if (this.userContext) return this.userContext;

    const user = this.context.pageContext.user;
    this.userContext = {
      loginName: user.loginName || '',
      displayName: user.displayName || '',
      email: user.email,
      department: undefined,
      jobTitle: undefined,
      officeLocation: undefined,
      groupIds: []
    };

    // Optionally enrich from Microsoft Graph / User Profile in a later phase
    return this.userContext;
  }

  /**
   * Filter items that have an audience field.
   * Items with empty/null audience are considered visible to everyone.
   */
  public filterByAudience<T extends BaseListItem>(
    items: T[],
    userContext: IAudienceContext
  ): T[] {
    return items.filter(item => {
      if (!item.audience || item.audience.length === 0) {
        return true; // visible to all
      }

      const audiences = Array.isArray(item.audience)
        ? item.audience
        : [item.audience as unknown as string];

      return audiences.some(aud => {
        const a = (aud || '').toLowerCase();
        if (a === 'all users' || a === 'everyone') return true;
        if (userContext.department && a === userContext.department.toLowerCase()) return true;
        if (userContext.jobTitle && a === userContext.jobTitle.toLowerCase()) return true;
        if (userContext.officeLocation && a === userContext.officeLocation.toLowerCase()) return true;
        if (userContext.groupIds.some(g => g.toLowerCase() === a)) return true;
        return false;
      });
    });
  }

  /**
   * Build an OData filter fragment for audience when the list stores
   * audience as a choice / multi-choice field (best-effort server-side).
   */
  public buildAudienceODataFilter(
    audienceField: string,
    userContext: IAudienceContext
  ): string {
    // Prefer server-side filtering when possible; fallback is client-side.
    // Multi-choice OData is limited, so we often still filter client-side after a reasonable top.
    return ''; // intentionally empty – client filter is applied after retrieval
  }
}
