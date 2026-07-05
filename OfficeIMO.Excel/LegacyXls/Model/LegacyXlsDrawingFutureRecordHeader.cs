namespace OfficeIMO.Excel.LegacyXls.Model {
    /// <summary>
    /// Describes the common future-record header used by preserve-only drawing stream records.
    /// </summary>
    public sealed class LegacyXlsDrawingFutureRecordHeader {
        /// <summary>
        /// Creates future-record header metadata.
        /// </summary>
        public LegacyXlsDrawingFutureRecordHeader(
            ushort wrappedRecordType,
            ushort flags,
            ushort? firstRow,
            ushort? lastRow,
            ushort? firstColumn,
            ushort? lastColumn,
            int streamByteCount) {
            if (streamByteCount < 0) {
                throw new ArgumentOutOfRangeException(nameof(streamByteCount));
            }

            WrappedRecordType = wrappedRecordType;
            Flags = flags;
            FirstRow = firstRow;
            LastRow = lastRow;
            FirstColumn = firstColumn;
            LastColumn = lastColumn;
            StreamByteCount = streamByteCount;
        }

        /// <summary>Gets the BIFF record type wrapped by the future-record stream.</summary>
        public ushort WrappedRecordType { get; }

        /// <summary>Gets the future-record option flags.</summary>
        public ushort Flags { get; }

        /// <summary>Gets whether the future-record header carries a cell range reference.</summary>
        public bool HasRange => (Flags & 0x0001) != 0;

        /// <summary>Gets whether all range fields declared by the future-record flags were decoded.</summary>
        public bool HasCompleteRangeReference =>
            !HasRange
            || (FirstRow.HasValue
                && LastRow.HasValue
                && FirstColumn.HasValue
                && LastColumn.HasValue);

        /// <summary>Gets the zero-based first row of the attached range, when present.</summary>
        public ushort? FirstRow { get; }

        /// <summary>Gets the zero-based last row of the attached range, when present.</summary>
        public ushort? LastRow { get; }

        /// <summary>Gets the zero-based first column of the attached range, when present.</summary>
        public ushort? FirstColumn { get; }

        /// <summary>Gets the zero-based last column of the attached range, when present.</summary>
        public ushort? LastColumn { get; }

        /// <summary>Gets the remaining future-record stream byte count after the decoded header.</summary>
        public int StreamByteCount { get; }
    }
}
