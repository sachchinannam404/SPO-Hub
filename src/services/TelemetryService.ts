/**
 * Shared Telemetry Service
 * Tracks usage without capturing sensitive information.
 */

import { Log } from '@microsoft/sp-core-library';
import { WebPartContext } from '@microsoft/sp-webpart-base';

const LOG_SOURCE = 'TelemetryService';

export interface TelemetryEvent {
  eventName: string;
  componentName: string;
  itemId?: string;
  pageUrl?: string;
  userContext?: string; // e.g. "authenticated" | "anonymous" – never PII
  timestamp: string;
  metadata?: Record<string, unknown>;
}

export type TelemetryEventName =
  | 'WebPartLoaded'
  | 'ItemViewed'
  | 'ItemClicked'
  | 'SearchPerformed'
  | 'FilterApplied'
  | 'LinkOpened'
  | 'ActionStarted'
  | 'ActionCompleted'
  | 'ActionFailed'
  | 'RequestSubmitted'
  | 'AutomationLaunched'
  | 'FeedbackSubmitted';

export class TelemetryService {
  private context: WebPartContext;
  private componentName: string;
  private enabled: boolean = true;

  constructor(context: WebPartContext, componentName: string) {
    this.context = context;
    this.componentName = componentName;
  }

  public track(
    eventName: TelemetryEventName | string,
    metadata?: Record<string, unknown>,
    itemId?: string
  ): void {
    if (!this.enabled) return;

    const event: TelemetryEvent = {
      eventName,
      componentName: this.componentName,
      itemId,
      pageUrl: this.context.pageContext?.web?.absoluteUrl,
      userContext: this.context.pageContext?.user?.loginName ? 'authenticated' : 'anonymous',
      timestamp: new Date().toISOString(),
      metadata: this.sanitize(metadata)
    };

    // Write to browser console / SPFx logger in development.
    // In production this can be forwarded to Application Insights / custom endpoint.
    Log.info(LOG_SOURCE, JSON.stringify(event));

    // Optional: push to window-level buffer for later flush
    if (typeof window !== 'undefined') {
      const buffer = (window as any).__spoHubTelemetry || [];
      buffer.push(event);
      (window as any).__spoHubTelemetry = buffer.slice(-100); // keep last 100
    }
  }

  public trackWebPartLoaded(metadata?: Record<string, unknown>): void {
    this.track('WebPartLoaded', metadata);
  }

  public trackItemClicked(itemId: string, metadata?: Record<string, unknown>): void {
    this.track('ItemClicked', metadata, itemId);
  }

  public trackSearch(query: string): void {
    this.track('SearchPerformed', { queryLength: query?.length || 0 });
  }

  public trackAction(action: string, success: boolean, metadata?: Record<string, unknown>): void {
    this.track(success ? 'ActionCompleted' : 'ActionFailed', {
      action,
      ...metadata
    });
  }

  private sanitize(metadata?: Record<string, unknown>): Record<string, unknown> | undefined {
    if (!metadata) return undefined;
    const clean: Record<string, unknown> = {};
    const blocked = ['email', 'password', 'token', 'secret', 'ssn', 'phone'];
    Object.keys(metadata).forEach(k => {
      const lower = k.toLowerCase();
      if (blocked.some(b => lower.indexOf(b) > -1)) return;
      clean[k] = metadata[k];
    });
    return clean;
  }
}
