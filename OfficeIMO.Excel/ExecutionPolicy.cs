namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls when worksheet mutations perform structural validation.
    /// </summary>
    public enum WorksheetValidationMode {
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
    /// Controls how heavy operations in OfficeIMO.Excel run (sequential vs parallel) based on workload size.
    /// Configure global and per‑operation thresholds and optionally observe decisions.
    /// </summary>
    public sealed class ExecutionPolicy {
        /// <summary>
        /// Global execution mode. When <see cref="ExecutionMode.Automatic"/>, the policy selects sequential or parallel per operation.
        /// </summary>
        public ExecutionMode Mode { get; set; } = ExecutionMode.Automatic;

        /// <summary>Default threshold above which Automatic switches to Parallel.</summary>
        public int ParallelThreshold { get; set; } = 10_000;

        /// <summary>Per-operation thresholds (names: "CellValues", "InsertObjects", "AutoFitColumns", ...).</summary>
        public Dictionary<string, int> OperationThresholds { get; } = new(StringComparer.Ordinal);

        /// <summary>Optional cap for parallel compute phase.</summary>
        public int? MaxDegreeOfParallelism { get; set; }

        /// <summary>Structured diagnostics (operation, items, decided mode).</summary>
        public Action<string, int, ExecutionMode>? OnDecision { get; set; }

        /// <summary>
        /// Optional timing callback invoked by long-running operations to report elapsed time.
        /// Provides a lightweight hook for performance monitoring in large workbooks.
        /// </summary>
        public Action<string, TimeSpan>? OnTiming { get; set; }

        /// <summary>
        /// Optional informational callback for verbose/debug diagnostics (no sheet output).
        /// Use to observe non-fatal events (e.g., grid overflow handled by Shrink/Summarize).
        /// </summary>
        public Action<string>? OnInfo { get; set; }

        /// <summary>
        /// Indicates whether consumers explicitly requested diagnostics. When true, operations configured with
        /// <see cref="WorksheetValidationMode.DiagnosticsOnly"/> will run validation even if no callbacks are wired.
        /// </summary>
        public bool DiagnosticsRequested { get; set; }

        /// <summary>
        /// Controls when worksheet mutation validation is executed. Defaults to running only when diagnostics
        /// are requested to avoid penalizing hot paths.
        /// </summary>
        public WorksheetValidationMode WorksheetValidation { get; set; } = WorksheetValidationMode.DiagnosticsOnly;

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

        internal bool AreDiagnosticsRequested
            => DiagnosticsRequested || OnInfo != null || OnTiming != null || OnDecision != null;

        /// <summary>
        /// Decide execution mode for a given operation and workload size.
        /// </summary>
        /// <param name="operationName">Descriptive operation name (e.g. "ReadRange", "AutoFitColumns").</param>
        /// <param name="itemCount">Approximate number of items to process.</param>
        internal ExecutionMode Decide(string operationName, int itemCount) {
            var thr = OperationThresholds.TryGetValue(operationName, out var v) ? v : ParallelThreshold;
            var decided = itemCount > thr ? ExecutionMode.Parallel : ExecutionMode.Sequential;
            OnDecision?.Invoke(operationName, itemCount, decided);
            return decided;
        }

        /// <summary>
        /// Creates a policy with recommended default thresholds for common operations.
        /// </summary>
        public ExecutionPolicy() {
            // Set recommended defaults
            OperationThresholds["CellValues"] = 10_000;
            OperationThresholds["InsertObjects"] = 1_000;
            OperationThresholds["AutoFitColumns"] = 2_000;
            OperationThresholds["AutoFitRows"] = 2_000;
            OperationThresholds["ConditionalFormatting"] = 2_000;
        }
    }
}
