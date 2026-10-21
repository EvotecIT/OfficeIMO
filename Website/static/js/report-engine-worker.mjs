import * as engine from "/assets/officeimo/officeimo.mjs";
import { handleWorkerExports } from "/assets/officeimo/officeimo-execution.mjs";
import { reportColumns, writeOperationalReport } from "/js/report-engine-job.mjs";
import { writeAnalyticalReport } from "/js/analytical-report-job.mjs";
import { writeWideAnalyticalReport } from "/js/wide-analytical-report-job.mjs";

// The host owns this Worker and its lifetime. The engine owns bounded messages and the writers.
handleWorkerExports(self, context => context.input?.layout === "wide"
  ? writeWideAnalyticalReport(engine, context)
  : context.input?.layout === "hierarchical"
  ? writeAnalyticalReport(engine, { ...context, columns: reportColumns })
  : writeOperationalReport(engine, context));
