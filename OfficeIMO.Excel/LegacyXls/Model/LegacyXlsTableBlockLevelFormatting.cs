namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Represents List12 block-level formatting references for projected table regions.
    /// </summary>
    public sealed class LegacyXlsTableBlockLevelFormatting {
        internal LegacyXlsTableBlockLevelFormatting(
            int? headerStyleRecordIndex,
            string? headerStyleName,
            int? dataStyleRecordIndex,
            string? dataStyleName,
            int? totalStyleRecordIndex,
            string? totalStyleName) {
            HeaderStyleRecordIndex = headerStyleRecordIndex;
            HeaderStyleName = headerStyleName;
            DataStyleRecordIndex = dataStyleRecordIndex;
            DataStyleName = dataStyleName;
            TotalStyleRecordIndex = totalStyleRecordIndex;
            TotalStyleName = totalStyleName;
        }

        /// <summary>Gets the zero-based Style record index for table header cells, when specified.</summary>
        public int? HeaderStyleRecordIndex { get; }

        /// <summary>Gets the header style name carried by the List12BlockLevel payload, when specified.</summary>
        public string? HeaderStyleName { get; }

        /// <summary>Gets the zero-based Style record index for table data cells, when specified.</summary>
        public int? DataStyleRecordIndex { get; }

        /// <summary>Gets the data style name carried by the List12BlockLevel payload, when specified.</summary>
        public string? DataStyleName { get; }

        /// <summary>Gets the zero-based Style record index for table total-row cells, when specified.</summary>
        public int? TotalStyleRecordIndex { get; }

        /// <summary>Gets the total-row style name carried by the List12BlockLevel payload, when specified.</summary>
        public string? TotalStyleName { get; }
    }
}
