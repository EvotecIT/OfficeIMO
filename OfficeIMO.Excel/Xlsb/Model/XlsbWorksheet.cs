namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbWorksheet {
        private readonly List<XlsbCell> _cells = new List<XlsbCell>();
        private readonly List<XlsbRowInfo> _rows = new List<XlsbRowInfo>();
        private readonly List<XlsbColumnInfo> _columns = new List<XlsbColumnInfo>();
        private readonly List<XlsbCellRange> _mergedRanges = new List<XlsbCellRange>();

        internal XlsbWorksheet(string name, string relationshipId, uint tabId, uint state) {
            Name = name;
            RelationshipId = relationshipId;
            TabId = tabId;
            State = state;
        }

        internal string Name { get; }

        internal string RelationshipId { get; }

        internal uint TabId { get; }

        internal uint State { get; }

        internal string? PartName { get; set; }

        internal IReadOnlyList<XlsbCell> Cells => _cells;

        internal IReadOnlyList<XlsbRowInfo> Rows => _rows;

        internal IReadOnlyList<XlsbColumnInfo> Columns => _columns;

        internal IReadOnlyList<XlsbCellRange> MergedRanges => _mergedRanges;

        internal XlsbCellRange? UsedRange { get; set; }

        internal XlsbWorksheetFormatInfo? FormatInfo { get; set; }

        internal XlsbPaneInfo? Pane { get; set; }

        internal void AddCell(XlsbCell cell) => _cells.Add(cell);

        internal void AddRow(XlsbRowInfo row) => _rows.Add(row);

        internal void AddColumn(XlsbColumnInfo column) => _columns.Add(column);

        internal void AddMergedRange(XlsbCellRange range) => _mergedRanges.Add(range);
    }
}
