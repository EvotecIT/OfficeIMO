import * as engine from "/assets/officeimo/officeimo.mjs";
import { handleWorkerExports } from "/assets/officeimo/officeimo-execution.mjs";
import { writeOperationalReport } from "/js/report-engine-job.mjs";

// The host owns this Worker and its lifetime. The engine owns bounded messages and the writers.
handleWorkerExports(self, context => writeOperationalReport(engine, context));
