/**
 * Worker entrypoint. Imports every capability module for its registration
 * side-effect, then exports the shared worker as the default export.
 */

import { worker } from "./worker.js";
import "./webhooks/oooRequestChanged.js"; // Notion automation → reconcile one row
import "./syncs/oooReconcileSweep.js"; // every 10m: re-reconcile + remove orphaned events
import "./syncs/holidaySync.js"; // every 6h: mirror the two holiday databases
import "./tools/setup.js"; // resolveO365Group / checkOooSetup / reconcileOooRequest

export default worker;
