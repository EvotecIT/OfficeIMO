namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbCell {
        internal XlsbCell(int row, int column, XlsbCellValueKind kind, object? value, uint styleIndex) {
            Row = row;
            Column = column;
            Kind = kind;
            Value = value;
            StyleIndex = styleIndex;
        }

        internal int Row { get; }

        internal int Column { get; }

        internal XlsbCellValueKind Kind { get; }

        internal object? Value { get; }

        internal uint StyleIndex { get; }

        internal string? FormulaText { get; set; }

        internal byte[]? FormulaBytes { get; set; }

        /// <summary>Gets the complete formula payload following the cached result, including flags and extra data.</summary>
        internal byte[]? FormulaPayloadBytes { get; set; }

        /// <summary>Gets the original BIFF12 cell record type for preservation-aware rewriting.</summary>
        internal int SourceRecordType { get; set; }

        /// <summary>Gets the original BIFF12 cell record payload for preservation-aware rewriting.</summary>
        internal byte[]? SourceRecordData { get; set; }
    }
}
