namespace OfficeIMO.Excel {
    /// <summary>
    /// Summary of workbook refresh-on-open metadata updates.
    /// </summary>
    public sealed class ExcelRefreshOnOpenResult {
        internal ExcelRefreshOnOpenResult(bool enabled, int pivotCacheCount, int connectionCount) {
            Enabled = enabled;
            PivotCacheCount = pivotCacheCount;
            ConnectionCount = connectionCount;
        }

        /// <summary>Whether refresh-on-open was enabled or disabled.</summary>
        public bool Enabled { get; }

        /// <summary>Number of pivot cache definitions updated.</summary>
        public int PivotCacheCount { get; }

        /// <summary>Number of workbook connection entries updated.</summary>
        public int ConnectionCount { get; }
    }
}
