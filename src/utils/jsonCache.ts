import fs from "fs";

require("dotenv").config();

const cacheDir = process.env.AUTH_CACHE_DIR;

if (!cacheDir) {
  throw new Error("AUTH_CACHE_DIR environment variable is not set");
}

type fileNames = "authentication-record" | "token-cache";

// Use error logs because it won't interfere with standard output, and it'll show up in the claude log

export function readJsonCache(fileName: fileNames) {
  const filePath = `${cacheDir}/${fileName}.json`;

  if (!fs.existsSync(filePath)) {
    return undefined;
  }

  try {
    const record = fs.readFileSync(filePath, "utf-8");
    console.error(`Loaded ${fileName}`);
    return record;
  } catch (error) {
    console.error(`Failed to load ${fileName}:`, error);
    // Optionally delete the corrupted file
    // fs.unlinkSync(authenticationRecordPath);
  }
}

export function writeJsonCache(fileName: fileNames, jsonStr: string) {
  if (!fs.existsSync(cacheDir!)) {
    fs.mkdirSync(cacheDir!, { recursive: true });
  }

  const filePath = `${cacheDir}/${fileName}.json`;

  fs.writeFileSync(filePath, jsonStr);
  console.error(`Saved ${fileName}`);
}
