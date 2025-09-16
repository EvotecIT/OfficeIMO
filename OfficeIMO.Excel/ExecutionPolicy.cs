using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel
{
    /// <summary>
    /// Controls how heavy operations in OfficeIMO.Excel run (sequential vs parallel) based on workload size.
    /// Configure global and per‑operation thresholds and optionally observe decisions.
    /// </summary>
    public sealed class ExecutionPolicy
    {
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
        /// Helper to invoke the timing callback if configured.
        /// </summary>
        internal void ReportTiming(string operation, TimeSpan elapsed)
            => OnTiming?.Invoke(operation, elapsed);

        internal void ReportInfo(string message)
            => OnInfo?.Invoke(message);

        /// <summary>
        /// Decide execution mode for a given operation and workload size.
        /// </summary>
        /// <param name="operationName">Descriptive operation name (e.g. "ReadRange", "AutoFitColumns").</param>
        /// <param name="itemCount">Approximate number of items to process.</param>
        internal ExecutionMode Decide(string operationName, int itemCount)
        {
            var thr = OperationThresholds.TryGetValue(operationName, out var v) ? v : ParallelThreshold;
            var decided = itemCount > thr ? ExecutionMode.Parallel : ExecutionMode.Sequential;
            OnDecision?.Invoke(operationName, itemCount, decided);
            return decided;
        }

        /// <summary>
        /// Creates a policy with recommended default thresholds for common operations.
        /// </summary>
        public ExecutionPolicy()
        {
            // Set recommended defaults
            OperationThresholds["CellValues"] = 10_000;
            OperationThresholds["InsertObjects"] = 1_000;
            OperationThresholds["AutoFitColumns"] = 2_000;
            OperationThresholds["AutoFitRows"] = 2_000;
            OperationThresholds["ConditionalFormatting"] = 2_000;
        }
    }
}
