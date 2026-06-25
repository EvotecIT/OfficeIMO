namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes a preserve-only DConRef record that points at a consolidation or PivotTable source range.
    /// </summary>
    public sealed class LegacyXlsDataConsolidationReference {
        /// <summary>
        /// Creates decoded DConRef metadata.
        /// </summary>
        public LegacyXlsDataConsolidationReference(
            int recordOffset,
            ushort recordType,
            int firstRow,
            int lastRow,
            int firstColumn,
            int lastColumn,
            string cellRange,
            LegacyXlsDataConsolidationSourceKind sourceKind,
            string source,
            byte? sourcePrefix,
            int unusedByteCount) {
            RecordOffset = recordOffset;
            RecordType = recordType;
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
            CellRange = cellRange ?? throw new ArgumentNullException(nameof(cellRange));
            SourceKind = sourceKind;
            Source = source ?? throw new ArgumentNullException(nameof(source));
            SourcePrefix = sourcePrefix;
            UnusedByteCount = unusedByteCount;
            RowSpan = LastRow >= FirstRow ? LastRow - FirstRow + 1 : 0;
            ColumnSpan = LastColumn >= FirstColumn ? LastColumn - FirstColumn + 1 : 0;
        }

        /// <summary>Gets the byte offset of the DConRef BIFF record.</summary>
        public int RecordOffset { get; }

        /// <summary>Gets the BIFF record type.</summary>
        public ushort RecordType { get; }

        /// <summary>Gets the first one-based row in the referenced source range.</summary>
        public int FirstRow { get; }

        /// <summary>Gets the last one-based row in the referenced source range.</summary>
        public int LastRow { get; }

        /// <summary>Gets the first one-based column in the referenced source range.</summary>
        public int FirstColumn { get; }

        /// <summary>Gets the last one-based column in the referenced source range.</summary>
        public int LastColumn { get; }

        /// <summary>Gets the referenced source range in A1 notation.</summary>
        public string CellRange { get; }

        /// <summary>Gets the number of rows covered by the source range, or zero when the range is invalid.</summary>
        public int RowSpan { get; }

        /// <summary>Gets the number of columns covered by the source range, or zero when the range is invalid.</summary>
        public int ColumnSpan { get; }

        /// <summary>Gets the decoded DConFile source kind.</summary>
        public LegacyXlsDataConsolidationSourceKind SourceKind { get; }

        /// <summary>Gets the workbook path or sheet name after the DConFile prefix.</summary>
        public string Source { get; }

        /// <summary>Gets the raw DConFile prefix byte when one was present.</summary>
        public byte? SourcePrefix { get; }

        /// <summary>Gets the count of unused trailing bytes after the DConFile string.</summary>
        public int UnusedByteCount { get; }
    }
}
