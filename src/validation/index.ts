import { ValidationError } from '../errors';

export function validateNotebookName(name: string): void {
  if (!name || name.trim().length === 0) {
    throw new ValidationError('Notebook name cannot be empty');
  }
  if (name.length > 128) {
    throw new ValidationError('Notebook name cannot exceed 128 characters');
  }
}

export function validateSectionName(name: string): void {
  if (!name || name.trim().length === 0) {
    throw new ValidationError('Section name cannot be empty');
  }
  if (name.length > 128) {
    throw new ValidationError('Section name cannot exceed 128 characters');
  }
}

export function validatePageTitle(title: string): void {
  if (!title || title.trim().length === 0) {
    throw new ValidationError('Page title cannot be empty');
  }
  if (title.length > 255) {
    throw new ValidationError('Page title cannot exceed 255 characters');
  }
}

export function validatePageContent(content: string): void {
  if (!content) {
    throw new ValidationError('Page content cannot be empty');
  }
  // Basic HTML validation
  if (!content.includes('<') || !content.includes('>')) {
    throw new ValidationError('Page content must be valid HTML');
  }
}

export function validateId(id: string, type: string): void {
  if (!id || id.trim().length === 0) {
    throw new ValidationError(`${type} ID cannot be empty`);
  }
  // OneNote IDs are typically 36 characters (UUID format)
  if (!/^[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}$/.test(id)) {
    throw new ValidationError(`Invalid ${type} ID format`);
  }
}