namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes a generated subtotal summary block.
    /// </summary>
    public sealed class ExcelSubtotalResult {
        internal ExcelSubtotalResult(
            string summaryRange,
            int summaryStartRow,
            int summaryEndRow,
            IReadOnlyList<ExcelSubtotalGroupResult> groups,
            bool grandTotalWritten) {
            SummaryRange = summaryRange;
            SummaryStartRow = summaryStartRow;
            SummaryEndRow = summaryEndRow;
            Groups = groups;
            GrandTotalWritten = grandTotalWritten;
        }

        /// <summary>A1 range occupied by the generated summary block.</summary>
        public string SummaryRange { get; }

        /// <summary>First row occupied by the generated summary block.</summary>
        public int SummaryStartRow { get; }

        /// <summary>Last row occupied by the generated summary block.</summary>
        public int SummaryEndRow { get; }

        /// <summary>Generated subtotal groups.</summary>
        public IReadOnlyList<ExcelSubtotalGroupResult> Groups { get; }

        /// <summary>Number of generated subtotal groups.</summary>
        public int GroupCount => Groups.Count;

        /// <summary>Whether a grand total row was written.</summary>
        public bool GrandTotalWritten { get; }
    }

    /// <summary>
    /// Describes a single generated subtotal group.
    /// </summary>
    public sealed class ExcelSubtotalGroupResult {
        internal ExcelSubtotalGroupResult(string key, int sourceStartRow, int sourceEndRow, int summaryRow) {
            Key = key;
            SourceStartRow = sourceStartRow;
            SourceEndRow = sourceEndRow;
            SummaryRow = summaryRow;
        }

        /// <summary>Group key read from the worksheet.</summary>
        public string Key { get; }

        /// <summary>First source data row included in the group.</summary>
        public int SourceStartRow { get; }

        /// <summary>Last source data row included in the group.</summary>
        public int SourceEndRow { get; }

        /// <summary>Generated summary row for the group.</summary>
        public int SummaryRow { get; }
    }
}
