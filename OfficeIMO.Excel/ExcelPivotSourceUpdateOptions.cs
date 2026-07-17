namespace OfficeIMO.Excel {
    /// <summary>
    /// Controls safety checks applied when an existing pivot table source is changed.
    /// </summary>
    public sealed class ExcelPivotSourceUpdateOptions {
        /// <summary>
        /// When true, the new source header names must match the existing pivot cache database fields.
        /// Field count must always remain stable because this operation does not rebuild pivot field definitions.
        /// </summary>
        public bool RequireMatchingHeaders { get; set; } = true;

        /// <summary>
        /// When true, a cache shared by multiple pivot tables may be updated. All affected pivots are returned in the result.
        /// </summary>
        public bool AllowSharedCacheUpdate { get; set; }
    }

    /// <summary>
    /// Describes a completed pivot table source update.
    /// </summary>
    public sealed class ExcelPivotSourceUpdateResult {
        internal ExcelPivotSourceUpdateResult(
            string pivotTableName,
            uint cacheId,
            string sourceSheet,
            string sourceRange,
            IReadOnlyList<string> affectedPivotTables,
            uint invalidatedCachedRecordCount) {
            PivotTableName = pivotTableName;
            CacheId = cacheId;
            SourceSheet = sourceSheet;
            SourceRange = sourceRange;
            AffectedPivotTables = affectedPivotTables;
            InvalidatedCachedRecordCount = invalidatedCachedRecordCount;
        }

        /// <summary>The requested pivot table name.</summary>
        public string PivotTableName { get; }

        /// <summary>The updated pivot cache identifier.</summary>
        public uint CacheId { get; }

        /// <summary>The new source worksheet name.</summary>
        public string SourceSheet { get; }

        /// <summary>The normalized new source range.</summary>
        public string SourceRange { get; }

        /// <summary>Pivot tables that reference the updated cache.</summary>
        public IReadOnlyList<string> AffectedPivotTables { get; }

        /// <summary>Number of stale cached source records removed from the package.</summary>
        public uint InvalidatedCachedRecordCount { get; }
    }
}
