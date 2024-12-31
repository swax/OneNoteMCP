export class OneNoteMCPError extends Error {
  constructor(message: string, public code: string, public statusCode?: number) {
    super(message);
    this.name = 'OneNoteMCPError';
  }
}

export class AuthenticationError extends OneNoteMCPError {
  constructor(message: string) {
    super(message, 'AUTHENTICATION_ERROR', 401);
  }
}

export class NotFoundError extends OneNoteMCPError {
  constructor(message: string) {
    super(message, 'NOT_FOUND_ERROR', 404);
  }
}

export class ValidationError extends OneNoteMCPError {
  constructor(message: string) {
    super(message, 'VALIDATION_ERROR', 400);
  }
}

export class RateLimitError extends OneNoteMCPError {
  constructor(message: string) {
    super(message, 'RATE_LIMIT_ERROR', 429);
  }
}

export class ServerError extends OneNoteMCPError {
  constructor(message: string) {
    super(message, 'SERVER_ERROR', 500);
  }
}