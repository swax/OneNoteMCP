import { NotebookManagement } from '../functions/notebooks';
import { TokenCredential, GetTokenOptions } from '@azure/identity';

// Mock TokenCredential
class MockTokenCredential implements TokenCredential {
  async getToken(scopes: string | string[], options?: GetTokenOptions) {
    return { token: 'mock-token', expiresOnTimestamp: Date.now() + 3600000 };
  }
}

// Mock Microsoft Graph Client
jest.mock('@microsoft/microsoft-graph-client', () => ({
  Client: {
    init: () => ({
      api: () => ({
        select: () => ({
          get: async () => ({
            value: [
              {
                id: 'test-id',
                displayName: 'Test Notebook',
                createdDateTime: '2024-01-01T00:00:00Z',
                lastModifiedDateTime: '2024-01-02T00:00:00Z',
                sectionsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/test-id/sections'
              }
            ]
          }),
          post: async () => ({
            id: 'new-test-id',
            displayName: 'New Test Notebook',
            createdDateTime: '2024-01-01T00:00:00Z',
            lastModifiedDateTime: '2024-01-01T00:00:00Z',
            sectionsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/new-test-id/sections'
          }),
          delete: async () => undefined
        })
      })
    })
  }
}));

describe('NotebookManagement', () => {
  let notebookManagement: NotebookManagement;

  beforeEach(() => {
    notebookManagement = new NotebookManagement(new MockTokenCredential());
  });

  describe('listNotebooks', () => {
    it('should return an array of notebooks', async () => {
      const notebooks = await notebookManagement.listNotebooks();
      expect(notebooks).toHaveLength(1);
      expect(notebooks[0]).toEqual({
        id: 'test-id',
        name: 'Test Notebook',
        createdTime: '2024-01-01T00:00:00Z',
        lastModifiedTime: '2024-01-02T00:00:00Z',
        sectionsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/test-id/sections'
      });
    });
  });

  describe('createNotebook', () => {
    it('should create a new notebook', async () => {
      const notebook = await notebookManagement.createNotebook({
        name: 'New Test Notebook'
      });
      expect(notebook).toEqual({
        id: 'new-test-id',
        name: 'New Test Notebook',
        createdTime: '2024-01-01T00:00:00Z',
        lastModifiedTime: '2024-01-01T00:00:00Z',
        sectionsUrl: 'https://graph.microsoft.com/v1.0/me/onenote/notebooks/new-test-id/sections'
      });
    });
  });

  describe('deleteNotebook', () => {
    it('should delete a notebook', async () => {
      await expect(notebookManagement.deleteNotebook({
        id: 'test-id'
      })).resolves.toBeUndefined();
    });
  });
});
