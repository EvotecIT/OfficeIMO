using System;
using System.Collections.Generic;

namespace OfficeIMO.Excel
{
    public sealed class ExecutionPolicy
    {
        public ExecutionMode Mode { get; set; } = ExecutionMode.Automatic;

        /// <summary>Default threshold above which Automatic switches to Parallel.</summary>
        public int ParallelThreshold { get; set; } = 10_000;

        /// <summary>Per-operation thresholds (names: "CellValues", "InsertObjects", "AutoFitColumns", ...).</summary>
        public Dictionary<string, int> OperationThresholds { get; } = new(StringComparer.Ordinal);

        /// <summary>Optional cap for parallel compute phase.</summary>
        public int? MaxDegreeOfParallelism { get; set; }

        /// <summary>Structured diagnostics (operation, items, decided mode).</summary>
        public Action<string, int, ExecutionMode>? OnDecision { get; set; }

        internal ExecutionMode Decide(string op, int count)
        {
            var thr = OperationThresholds.TryGetValue(op, out var v) ? v : ParallelThreshold;
            var decided = count > thr ? ExecutionMode.Parallel : ExecutionMode.Sequential;
            OnDecision?.Invoke(op, count, decided);
            return decided;
        }

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