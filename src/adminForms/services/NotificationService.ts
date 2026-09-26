/**
 * Power Automate / notification bridge for Admin Forms status changes.
 */

export interface INotificationConfig {
  powerAutomateWebhookUrl?: string;
}

export class NotificationService {
  private static config: INotificationConfig = {};

  public static configure(config: INotificationConfig): void {
    this.config = { ...this.config, ...config };
  }

  public static async notifyStatusChange(
    request: any,
    previousStatus: string,
    actorName: string
  ): Promise<void> {
    const url = this.config.powerAutomateWebhookUrl;
    if (!url) return;
    try {
      await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          event: 'StatusChange',
          previousStatus,
          actorName,
          request,
          timestamp: new Date().toISOString()
        })
      });
    } catch (e) {
      console.warn('[NotificationService] notifyStatusChange failed', e);
    }
  }

  public static async notifyBulkStatusChange(
    ids: string[],
    status: string,
    actorName: string
  ): Promise<void> {
    const url = this.config.powerAutomateWebhookUrl;
    if (!url) return;
    try {
      await fetch(url, {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
          event: 'BulkStatusChange',
          ids,
          status,
          actorName,
          timestamp: new Date().toISOString()
        })
      });
    } catch (e) {
      console.warn('[NotificationService] notifyBulkStatusChange failed', e);
    }
  }
}
