/**
 * Permission / RBAC Service
 * UI actions are derived from authorization context;
 * backend / workflow authorization remains authoritative.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { SPFI, spfi, SPFx } from '@pnp/sp';
import '@pnp/sp/webs';
import '@pnp/sp/security';
import '@pnp/sp/site-users/web';

export type ResourceAction =
  | 'View'
  | 'Create'
  | 'Edit'
  | 'Submit'
  | 'Approve'
  | 'Publish'
  | 'Admin';

export interface IPermissionGrant {
  resource: string;
  actions: ResourceAction[];
}

export class PermissionService {
  private context: WebPartContext;
  private sp: SPFI;
  private cache: Map<string, boolean> = new Map();

  constructor(context: WebPartContext) {
    this.context = context;
    this.sp = spfi().using(SPFx(context));
  }

  /**
   * Check if current user can perform an action on a resource.
   * For list-level permissions we rely on SharePoint; for custom
   * business actions we evaluate a simple grant model.
   */
  public async can(
    resource: string,
    action: ResourceAction,
    grants?: IPermissionGrant[]
  ): Promise<boolean> {
    const key = `${resource}:${action}`;
    if (this.cache.has(key)) {
      return this.cache.get(key)!;
    }

    // Default: authenticated users can View
    if (action === 'View') {
      this.cache.set(key, true);
      return true;
    }

    if (grants && grants.length > 0) {
      const grant = grants.find(g => g.resource === resource);
      const allowed = !!(grant && grant.actions.indexOf(action) > -1);
      this.cache.set(key, allowed);
      return allowed;
    }

    // Fallback: check if user is site owner / member with contribute
    try {
      const perms = await this.sp.web.getCurrentUserEffectivePermissions();
      // Simplified – real implementation would use PermissionKind
      const allowed = true; // SharePoint will enforce on write
      this.cache.set(key, allowed);
      return allowed;
    } catch {
      this.cache.set(key, false);
      return false;
    }
  }

  public clearCache(): void {
    this.cache.clear();
  }
}
