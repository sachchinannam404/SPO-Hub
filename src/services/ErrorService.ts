/**
 * Centralized Error Handling
 */

export enum ErrorCategory {
  ConfigurationError = 'ConfigurationError',
  PermissionError = 'PermissionError',
  DataError = 'DataError',
  NetworkError = 'NetworkError',
  AutomationError = 'AutomationError',
  ValidationError = 'ValidationError',
  UnknownError = 'UnknownError'
}

export class AppError extends Error {
  public readonly category: ErrorCategory;
  public readonly technicalDetails?: any;
  public readonly userMessage: string;

  constructor(category: ErrorCategory, userMessage: string, technicalDetails?: any) {
    super(userMessage);
    this.name = 'AppError';
    this.category = category;
    this.userMessage = userMessage;
    this.technicalDetails = technicalDetails;
  }
}

export function getUserFacingMessage(error: unknown): string {
  if (error instanceof AppError) {
    return error.userMessage;
  }
  return 'An unexpected error occurred. Please try again or contact your administrator.';
}

export function isConfigurationError(error: unknown): boolean {
  return error instanceof AppError && error.category === ErrorCategory.ConfigurationError;
}

export function isPermissionError(error: unknown): boolean {
  return error instanceof AppError && error.category === ErrorCategory.PermissionError;
}
