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
    }
}
