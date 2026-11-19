using OfficeIMO.Excel.Xlsb.Package;

namespace OfficeIMO.Excel.Xlsb.Model {
    /// <summary>A workbook sheet bundle and its decoded worksheet or chart-sheet content.</summary>
    internal sealed class XlsbSheet {
        private readonly List<XlsbCell> _cells = new List<XlsbCell>();
        private readonly List<XlsbRowInfo> _rows = new List<XlsbRowInfo>();
        private readonly List<XlsbColumnInfo> _columns = new List<XlsbColumnInfo>();
        private readonly List<XlsbCellRange> _mergedRanges = new List<XlsbCellRange>();
        private readonly List<XlsbHyperlink> _hyperlinks = new List<XlsbHyperlink>();
        private readonly List<DocumentFormat.OpenXml.OpenXmlElement> _featureElements = new List<DocumentFormat.OpenXml.OpenXmlElement>();

        internal XlsbSheet(string name, string relationshipId, uint tabId, uint state) {
            Name = name;
            RelationshipId = relationshipId;
            TabId = tabId;
            State = state;
        }

        internal bool IsChartSheet { get; set; }

        internal string? ChartCodeName { get; set; }

        internal string? DrawingRelationshipId { get; set; }

        /// <summary>The relationship id BrtLegacyDrawing stores when it names the VML drawing of the worksheet's comments.</summary>
        internal string? LegacyDrawingRelationshipId { get; set; }

        /// <summary>The worksheet's legacy comments, or null when it has no Comments part.</summary>
        internal XlsbComments? Comments { get; set; }

        /// <summary>The package path of the drawing part a worksheet's BrtDrawing links, once its relationship is resolved.</summary>
        internal string? DrawingPartName { get; set; }

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

        /// <summary>The primary window's view settings and selections, when the sheet has a readable view.</summary>
        internal XlsbSheetView? View { get; set; }

        /// <summary>Whether the primary window's sheet view selects this tab (fSelected), when the sheet has a view.</summary>
        internal bool? TabSelected { get; set; }

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

        /// <summary>
        /// Worksheet children decoded from BIFF12 record blocks that project directly to SpreadsheetML, such as page
        /// breaks or ignored-error ranges. The rewrite compares the document model against these to keep unchanged records.
        /// </summary>
        internal IReadOnlyList<DocumentFormat.OpenXml.OpenXmlElement> FeatureElements => _featureElements;

        internal void AddFeatureElement(DocumentFormat.OpenXml.OpenXmlElement element) => _featureElements.Add(element);

        /// <summary>
        /// True when the source holds data validations that could not be decoded, so none were projected and the native
        /// records may only be copied unchanged.
        /// </summary>
        internal bool HasUndecodedDataValidations { get; set; }

        /// <summary>
        /// True when the worksheet has conditional formatting that is kept as native records without a projection, such as
        /// Office 2010 rule extensions; replacing the projected collections would then lose or orphan that content.
        /// </summary>
        internal bool HasUnprojectedConditionalFormatting { get; set; }

        /// <summary>Tables decoded from the worksheet's Table parts, keyed by their worksheet relationship id.</summary>
        internal Dictionary<string, XlsbTable> Tables { get; } = new Dictionary<string, XlsbTable>(StringComparer.Ordinal);
    }
}
