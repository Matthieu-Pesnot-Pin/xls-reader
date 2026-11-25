/**
 * Ce que `--migrate-to-manager` sait de xls-reader (cf. `manager-bundle.ts`) :
 * le MCP n'a ni dossier de stockage ni variable propre, seule sa configuration
 * est exportée.
 */
import { APP_VERSION } from "../version.js";
import { runManagerExport, type ManagerExportSpec } from "./manager-bundle.js";

const spec: ManagerExportSpec = {
  packageName: "@imenam/xls-reader",
  packageVersion: APP_VERSION,
};

export function runMigrateToManager(argv: string[], cwd: string): Promise<number> {
  return runManagerExport(argv, cwd, spec);
}
