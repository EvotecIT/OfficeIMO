namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents the workbook sheet tab identifier array parsed from a BIFF TabId record.
    /// </summary>
    public sealed class LegacyXlsSheetTabIdCollection {
        /// <summary>
        /// Creates a sheet tab identifier collection.
        /// </summary>
        /// <param name="tabIds">Sheet tab identifiers in workbook order.</param>
        public LegacyXlsSheetTabIdCollection(IReadOnlyList<ushort> tabIds) {
            TabIds = new List<ushort>(tabIds).AsReadOnly();
        }

        /// <summary>Gets sheet tab identifiers in workbook order.</summary>
        public IReadOnlyList<ushort> TabIds { get; }
    }
}
