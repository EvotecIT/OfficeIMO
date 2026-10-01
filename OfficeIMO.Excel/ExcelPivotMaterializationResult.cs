using System.Collections.Generic;

namespace OfficeIMO.Excel {
    /// <summary>Describes pivot views generated from worksheet values without an Office refresh.</summary>
    public sealed class ExcelPivotMaterializationResult {
        internal ExcelPivotMaterializationResult(string name, string range, int records, ExcelMutationResult mutation,
            IReadOnlyList<string> affectedPivotTables) {
            PivotTableName = name;
            OutputRange = range;
            SourceRecordCount = records;
            Mutation = mutation;
            AffectedPivotTables = affectedPivotTables;
        }
        /// <summary>Name of the generated pivot view.</summary>
        public string PivotTableName { get; }
        /// <summary>Worksheet range containing the generated headers, values and totals.</summary>
        public string OutputRange { get; }
        /// <summary>Number of worksheet data records used, excluding the header.</summary>
        public int SourceRecordCount { get; }
        /// <summary>Transactional mutation result, including package validation diagnostics.</summary>
        public ExcelMutationResult Mutation { get; }
        /// <summary>Pivot views refreshed together because they use the same cache.</summary>
        public IReadOnlyList<string> AffectedPivotTables { get; }
    }
}
