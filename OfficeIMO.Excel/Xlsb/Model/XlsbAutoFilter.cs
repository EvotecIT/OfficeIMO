namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>Represents a worksheet AutoFilter and its common equality-list criteria.</summary>
    internal sealed class XlsbAutoFilter {
        private readonly List<XlsbAutoFilterColumn> _columns = new List<XlsbAutoFilterColumn>();

        internal XlsbAutoFilter(XlsbCellRange range) {
            Range = range ?? throw new ArgumentNullException(nameof(range));
        }

        internal XlsbCellRange Range { get; }

        internal IReadOnlyList<XlsbAutoFilterColumn> Columns => _columns;

        internal bool HasUnsupportedContent { get; set; }

        internal void AddColumn(XlsbAutoFilterColumn column) => _columns.Add(column);
    }

    /// <summary>Represents one column in a worksheet AutoFilter.</summary>
    internal sealed class XlsbAutoFilterColumn {
        private readonly List<string> _values = new List<string>();

        internal XlsbAutoFilterColumn(uint columnId) {
            ColumnId = columnId;
        }

        internal uint ColumnId { get; }

        internal bool IncludeBlank { get; set; }

        internal IReadOnlyList<string> Values => _values;

        internal bool HasUnsupportedContent { get; set; }

        internal void AddValue(string value) => _values.Add(value);
    }
}
