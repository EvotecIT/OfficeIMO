namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls when worksheet mutations perform structural validation.
    /// </summary>
    public enum ExcelWorksheetValidationMode {
        /// <summary>No validation is performed after write operations.</summary>
        Disabled,

        /// <summary>
        /// Run validation only when diagnostics are requested via callbacks or explicit opt-in.
        /// </summary>
        DiagnosticsOnly,

        /// <summary>
        /// Run lightweight validation in all builds and enable full Open XML validation in Debug builds only.
        /// </summary>
        DebugOnly,

        /// <summary>Always run structural validation regardless of diagnostics configuration.</summary>
        Always,
    }

    /// <summary>
    /// Controls how eligible compute work in OfficeIMO.Excel runs (sequential vs parallel) based on workload size.
    /// Specialized single-pass readers remain eligible in every mode because bypassing a faster reader does not
    /// make an operation meaningfully more parallel. Configure global and per-operation thresholds and optionally
    /// observe the execution strategy that was actually selected.
    /// </summary>
    public sealed class ExcelExecutionPolicy {
        /// <summary>
        /// Global execution mode. When <see cref="ExcelExecutionMode.Automatic"/>, the policy selects sequential or parallel per operation.
        /// </summary>
        public ExcelExecutionMode Mode { get; set; } = ExcelExecutionMode.Automatic;

        /// <summary>Default threshold above which Automatic permits parallel compute when no faster specialized reader applies.</summary>
        public int ParallelThreshold { get; set; } = 10_000;

        /// <summary>Per-operation thresholds (names: "CellValues", "InsertObjects", "InsertObjects.PowerShellProjection", "AutoFitColumns", ...).</summary>
        public Dictionary<string, int> OperationThresholds { get; } = new(StringComparer.Ordinal);

        /// <summary>Optional cap for parallel compute phase.</summary>
        public int? MaxDegreeOfParallelism { get; set; }

        /// <summary>Structured diagnostics (operation, items, decided mode).</summary>
        public Action<string, int, ExcelExecutionMode>? OnDecision { get; set; }

        /// <summary>
        /// Optional timing callback invoked by long-running operations to report elapsed time.
        /// Provides a lightweight hook for performance monitoring in large workbooks. Operation
        /// names may include whole operations such as "AutoFitColumns" and scoped sub-stages such
        /// as "AutoFitColumns.BuildPlan", "AutoFitColumns.CalculateWidths", or "AutoFitColumns.ApplyWidths".
        /// </summary>
        public Action<string, TimeSpan>? OnTiming { get; set; }

        /// <summary>
        /// Optional informational callback for verbose/debug diagnostics (no sheet output).
        /// Use to observe non-fatal events (e.g., grid overflow handled by Shrink/Summarize).
        /// </summary>
        public Action<string>? OnInfo { get; set; }

        /// <summary>
        /// Indicates whether consumers explicitly requested diagnostics. When true, operations configured with
        /// <see cref="ExcelWorksheetValidationMode.DiagnosticsOnly"/> will run validation even if no callbacks are wired.
        /// </summary>
        public bool DiagnosticsRequested { get; set; }

        /// <summary>
        /// Controls when worksheet mutation validation is executed. Defaults to running only when diagnostics
        /// are requested to avoid penalizing hot paths.
        /// </summary>
        public ExcelWorksheetValidationMode WorksheetValidation { get; set; } = ExcelWorksheetValidationMode.DiagnosticsOnly;

        /// <summary>
        /// Saves the worksheet part immediately after AutoFit mutates row heights or column widths.
        /// Disable for large report-generation pipelines that call <see cref="ExcelDocument.Save()"/> or dispose
        /// the document after a batch of worksheet mutations.
        /// </summary>
        public bool SaveWorksheetAfterAutoFit { get; set; } = true;

        /// <summary>
        /// Enables invoking <see cref="DocumentFormat.OpenXml.Validation.OpenXmlValidator"/> while debugging. This
        /// incurs a significant cost and is ignored when not compiling in <c>DEBUG</c> mode.
        /// </summary>
        public bool UseOpenXmlValidatorInDebug { get; set; } = true;

        /// <summary>
        /// Helper to invoke the timing callback if configured.
        /// </summary>
        internal void ReportTiming(string operation, TimeSpan elapsed)
            => OnTiming?.Invoke(operation, elapsed);

        internal void ReportInfo(string message)
            => OnInfo?.Invoke(message);

        internal void ReportDecision(string operationName, int itemCount, ExcelExecutionMode mode)
            => OnDecision?.Invoke(operationName, itemCount, mode);

        internal bool AreDiagnosticsRequested
            => DiagnosticsRequested || OnInfo != null || OnTiming != null || OnDecision != null;

        /// <summary>
        /// Decide execution mode for a given operation and workload size.
        /// </summary>
        /// <param name="operationName">Descriptive operation name (e.g. "ReadRange", "AutoFitColumns").</param>
        /// <param name="itemCount">Approximate number of items to process.</param>
        internal ExcelExecutionMode Decide(string operationName, int itemCount) {
            var thr = OperationThresholds.TryGetValue(operationName, out var v) ? v : ParallelThreshold;
            var decided = itemCount > thr ? ExcelExecutionMode.Parallel : ExcelExecutionMode.Sequential;
            ReportDecision(operationName, itemCount, decided);
            return decided;
        }

        /// <summary>
        /// Creates a policy with recommended default thresholds for common operations.
        /// </summary>
        public ExcelExecutionPolicy() {
            // Set recommended defaults
            OperationThresholds["CellValues"] = 10_000;
            OperationThresholds["InsertObjects"] = 1_000;
            OperationThresholds["InsertObjects.PowerShellProjection"] = 10_000;
            OperationThresholds["AutoFitColumns"] = 2_000;
            OperationThresholds["AutoFitRows"] = 2_000;
            OperationThresholds["ConditionalFormatting"] = 2_000;
        }
    }
}
