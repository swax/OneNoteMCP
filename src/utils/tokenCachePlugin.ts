/**
 * This works around an issue with the standard MSAL Node cache plugin
 * import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
 *
 * Where when Claude tries to run it, node:path module is unable to be resolved
 */

import { IdentityPlugin, TokenCachePersistenceOptions } from "@azure/identity";
import { ICachePlugin, TokenCacheContext } from "@azure/msal-node";
import { readJsonCache, writeJsonCache } from "./jsonCache";

interface CachePluginControl {
  setPersistence(
    persistenceFactory: (
      options?: TokenCachePersistenceOptions,
    ) => Promise<ICachePlugin>,
  ): void;
}

interface AzurePluginContext {
  cachePluginControl: CachePluginControl;
}

const createPersistenceCachePlugin = async (
  _options?: TokenCachePersistenceOptions,
): Promise<ICachePlugin> => {
  return {
    async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      const tokenCache = readJsonCache("token-cache");
      if (tokenCache) {
        cacheContext.tokenCache.deserialize(tokenCache);
      }
    },
    async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      // Save cache data if it has changed
      if (cacheContext.cacheHasChanged) {
        writeJsonCache("token-cache", cacheContext.tokenCache.serialize());
      }
      await Promise.resolve();
    },
  };
};

export const cachePersistencePlugin: IdentityPlugin = (context) => {
  const { cachePluginControl } = context as AzurePluginContext;

  cachePluginControl.setPersistence(createPersistenceCachePlugin);
};
