namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbWorkbook {
        private readonly List<XlsbWorksheet> _worksheets = new List<XlsbWorksheet>();
        private readonly List<XlsbImportDiagnostic> _diagnostics = new List<XlsbImportDiagnostic>();
        private readonly List<XlsbPreservedRecordInfo> _preservedRecords = new List<XlsbPreservedRecordInfo>();

        internal XlsbWorkbook(byte[] originalPackageBytes) {
            OriginalPackageBytes = originalPackageBytes ?? throw new ArgumentNullException(nameof(originalPackageBytes));
        }

        internal byte[] OriginalPackageBytes { get; }

        internal IReadOnlyList<XlsbWorksheet> Worksheets => _worksheets;

        internal IReadOnlyList<XlsbImportDiagnostic> Diagnostics => _diagnostics;

        internal IReadOnlyList<XlsbPreservedRecordInfo> PreservedRecords => _preservedRecords;

        internal int SharedStringCount { get; set; }

        internal bool Uses1904DateSystem { get; set; }

        internal XlsbCalculationProperties? CalculationProperties { get; set; }

        internal XlsbWorkbookProtection? WorkbookProtection { get; set; }

        internal XlsbStylesheet? Stylesheet { get; set; }

        internal void AddWorksheet(XlsbWorksheet worksheet) => _worksheets.Add(worksheet);

        internal void AddDiagnostic(XlsbImportDiagnostic diagnostic) => _diagnostics.Add(diagnostic);

        internal void AddPreservedRecord(XlsbPreservedRecordInfo record) => _preservedRecords.Add(record);
    }
}
