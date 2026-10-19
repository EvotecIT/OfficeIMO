using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>A workbook sheet bundle and its decoded worksheet or chart-sheet content.</summary>
    internal sealed class XlsbSheet {
        private readonly List<XlsbCell> _cells = new List<XlsbCell>();
        private readonly List<XlsbRowInfo> _rows = new List<XlsbRowInfo>();
        private readonly List<XlsbColumnInfo> _columns = new List<XlsbColumnInfo>();
        private readonly List<XlsbCellRange> _mergedRanges = new List<XlsbCellRange>();
        private readonly List<XlsbHyperlink> _hyperlinks = new List<XlsbHyperlink>();

        internal XlsbSheet(string name, string relationshipId, uint tabId, uint state) {
            Name = name;
            RelationshipId = relationshipId;
            TabId = tabId;
            State = state;
        }

        internal bool IsChartSheet { get; set; }

        internal string? ChartCodeName { get; set; }

        internal string? DrawingRelationshipId { get; set; }

        internal string Name { get; }

        internal string RelationshipId { get; }

        internal uint TabId { get; }

        internal uint State { get; }

        internal string? PartName { get; set; }

        internal IReadOnlyList<OfficeIMO.Excel.Xlsb.Biff12.XlsbFormulaGroup> FormulaGroups { get; set; } = Array.Empty<OfficeIMO.Excel.Xlsb.Biff12.XlsbFormulaGroup>();

        internal IReadOnlyList<XlsbCell> Cells => _cells;

        internal IReadOnlyList<XlsbRowInfo> Rows => _rows;

        internal IReadOnlyList<XlsbColumnInfo> Columns => _columns;

        internal IReadOnlyList<XlsbCellRange> MergedRanges => _mergedRanges;

        internal IReadOnlyList<XlsbHyperlink> Hyperlinks => _hyperlinks;

        internal IReadOnlyDictionary<string, XlsbPackageRelationship> Relationships { get; set; } =
            new Dictionary<string, XlsbPackageRelationship>(StringComparer.Ordinal);

        internal XlsbCellRange? UsedRange { get; set; }

        internal XlsbWorksheetProperties? Properties { get; set; }

        internal XlsbWorksheetFormatInfo? FormatInfo { get; set; }

        internal XlsbPaneInfo? Pane { get; set; }

        internal bool? ShowGridLines { get; set; }

        internal XlsbWorksheetProtection? Protection { get; set; }

        internal XlsbAutoFilter? AutoFilter { get; set; }

        internal XlsbPrintOptions? PrintOptions { get; set; }

        internal XlsbPageMargins? PageMargins { get; set; }

        internal XlsbPageSetup? PageSetup { get; set; }

        internal XlsbHeaderFooter? HeaderFooter { get; set; }

        internal void AddCell(XlsbCell cell) => _cells.Add(cell);

        internal void AddRow(XlsbRowInfo row) => _rows.Add(row);

        internal void AddColumn(XlsbColumnInfo column) => _columns.Add(column);

        internal void AddMergedRange(XlsbCellRange range) => _mergedRanges.Add(range);

        internal void AddHyperlink(XlsbHyperlink hyperlink) => _hyperlinks.Add(hyperlink);
    }
}
