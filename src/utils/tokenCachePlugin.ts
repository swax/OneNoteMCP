/**
 * This works around an issue with the standard MSAL Node cache plugin
 * import { cachePersistencePlugin } from "@azure/identity-cache-persistence";
 *
 * Where when Claude tries to run it, node:path module is unable to be resolved
 */

import { IdentityPlugin, TokenCachePersistenceOptions } from "@azure/identity";
import { ICachePlugin, TokenCacheContext } from "@azure/msal-node";
import fs from "fs";

const tokenCachePath = ".cache/token-cache.json";

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

// Mock implementation for createPersistenceCachePlugin
const createPersistenceCachePlugin = async (
  _options?: TokenCachePersistenceOptions,
): Promise<ICachePlugin> => {
  return {
    async beforeCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      if (fs.existsSync(tokenCachePath)) {
        try {
          const cacheData = fs.readFileSync(tokenCachePath, "utf-8");
          cacheContext.tokenCache.deserialize(cacheData);
          console.log("Loaded token cache");
        } catch (error) {
          console.error("Failed to load token cache:", error);
          // Optionally delete the corrupted file
          // fs.unlinkSync(authenticationRecordPath);
        }
      }
      await Promise.resolve();
    },
    async afterCacheAccess(cacheContext: TokenCacheContext): Promise<void> {
      // Save cache data if it has changed
      if (cacheContext.cacheHasChanged) {
        if (!fs.existsSync(".cache")) {
          fs.mkdirSync(".cache", { recursive: true });
        }

        fs.writeFileSync(tokenCachePath, cacheContext.tokenCache.serialize());
        console.log("Saved token cache");
      }
      await Promise.resolve();
    },
  };
};

export const cachePersistencePlugin: IdentityPlugin = (context) => {
  const { cachePluginControl } = context as AzurePluginContext;

  cachePluginControl.setPersistence(createPersistenceCachePlugin);
};
