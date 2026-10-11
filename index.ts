/**
 * The Outlook extension: what prifly imports. `activate` is given the api and
 * registers the MCP tools; the launcher exports join it in part 2. Types come
 * from `@prifly/extension-api` with `import type`, so nothing of it is loaded
 * at run time.
 */

import type { ExtensionApi } from "@prifly/extension-api";
import { outlookTools } from "./tools";

let api: ExtensionApi | null = null;

export function activate(given: ExtensionApi): () => void {
  api = given;
  // Absent on an older prifly: optional, like every part added after the first version.
  given.tools?.register(outlookTools(() => api));
  given.log("activated");
  return () => {
    api = null;
  };
}
