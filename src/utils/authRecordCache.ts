import { AuthenticationRecord } from "@azure/identity";
import fs from "fs";

const authenticationRecordPath = ".cache/authentication-record.json";

export function readAuthRecordCache() {
  if (fs.existsSync(authenticationRecordPath)) {
    try {
      const record = JSON.parse(
        fs.readFileSync(authenticationRecordPath, "utf-8"),
      );
      console.log("Loaded authentication record");
      return record;
    } catch (error) {
      console.error("Failed to load authentication record:", error);
      // Optionally delete the corrupted file
      // fs.unlinkSync(authenticationRecordPath);
    }
  }
  return undefined;
}

export function writeAuthRecordCache(record: AuthenticationRecord) {
  if (!fs.existsSync(".cache")) {
    fs.mkdirSync(".cache", { recursive: true });
  }

  fs.writeFileSync(authenticationRecordPath, JSON.stringify(record));
  console.log("Saved authentication record");
}
