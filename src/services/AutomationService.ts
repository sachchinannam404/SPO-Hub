/**
 * Automation Integration Pattern
 * Web parts never contain business-specific Power Automate logic.
 * They call this service with a configured flow endpoint.
 */

import { WebPartContext } from '@microsoft/sp-webpart-base';
import { AppError, ErrorCategory } from './ErrorService';
import { TelemetryService } from './TelemetryService';

export interface IFlowTriggerPayload {
  [key: string]: any;
}

export interface IFlowTriggerResult {
  success: boolean;
  statusCode?: number;
  message?: string;
  correlationId?: string;
}

export class AutomationService {
  private context: WebPartContext;
  private telemetry: TelemetryService;

  constructor(context: WebPartContext, telemetry: TelemetryService) {
    this.context = context;
    this.telemetry = telemetry;
  }

  /**
   * Trigger a Power Automate HTTP request flow.
   * The flow URL must be an approved, non-secret endpoint
   * (e.g. Power Automate HTTP trigger with Azure AD auth or shared access).
   */
  public async triggerFlow(
    flowUrl: string,
    payload: IFlowTriggerPayload,
    actionLabel?: string
  ): Promise<IFlowTriggerResult> {
    if (!flowUrl || !flowUrl.startsWith('https://')) {
      throw new AppError(
        ErrorCategory.ConfigurationError,
        'Automation flow URL is not configured correctly.'
      );
    }

    this.telemetry.track('ActionStarted', {
      action: actionLabel || 'TriggerFlow'
    });

    try {
      const body = {
        ...payload,
        _context: {
          siteUrl: this.context.pageContext.web.absoluteUrl,
          userLogin: this.context.pageContext.user.loginName,
          userDisplayName: this.context.pageContext.user.displayName,
          timestamp: new Date().toISOString()
        }
      };

      const response = await fetch(flowUrl, {
        method: 'POST',
        headers: {
          'Content-Type': 'application/json',
          Accept: 'application/json'
        },
        body: JSON.stringify(body)
      });

      const result: IFlowTriggerResult = {
        success: response.ok,
        statusCode: response.status,
        message: response.ok ? 'Request submitted successfully.' : 'Request failed.',
        correlationId: response.headers.get('x-ms-correlation-id') || undefined
      };

      this.telemetry.trackAction(actionLabel || 'TriggerFlow', result.success, {
        statusCode: result.statusCode
      });

      if (!result.success) {
        throw new AppError(
          ErrorCategory.AutomationError,
          'The automation request could not be completed. Please try again later.',
          { status: response.status }
        );
      }

      return result;
    } catch (error) {
      if (error instanceof AppError) throw error;
      this.telemetry.trackAction(actionLabel || 'TriggerFlow', false);
      throw new AppError(
        ErrorCategory.AutomationError,
        'Unable to reach the automation service.',
        error
      );
    }
  }
}
