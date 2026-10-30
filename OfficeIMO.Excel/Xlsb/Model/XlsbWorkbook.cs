namespace OfficeIMO.Excel.Xlsb.Model {
    internal sealed class XlsbWorkbook {
        private readonly List<XlsbSheet> _worksheets = new List<XlsbSheet>();
        private readonly List<XlsbDefinedName> _definedNames = new List<XlsbDefinedName>();
        private Dictionary<uint, HashSet<string>>? _localNamesBySheet;
        private HashSet<string>? _allLocalNames;
        private readonly List<XlsbExternalSheetReference> _externalSheetReferences = new List<XlsbExternalSheetReference>();
        private readonly List<bool> _selfSupportingLinks = new List<bool>();
        private readonly List<XlsbImportDiagnostic> _diagnostics = new List<XlsbImportDiagnostic>();
        private readonly List<XlsbPreservedRecordInfo> _preservedRecords = new List<XlsbPreservedRecordInfo>();

        internal XlsbWorkbook(byte[] originalPackageBytes, int maxPartBytes, long maxPackageBytes) {
            OriginalPackageBytes = originalPackageBytes ?? throw new ArgumentNullException(nameof(originalPackageBytes));
            MaxPartBytes = maxPartBytes;
            MaxPackageBytes = maxPackageBytes;
        }

        internal byte[] OriginalPackageBytes { get; }

        internal int MaxPartBytes { get; }

        internal long MaxPackageBytes { get; }

        internal IReadOnlyList<XlsbSheet> Sheets => _worksheets;

        internal IReadOnlyList<XlsbSheet> Worksheets => _worksheets.Where(sheet => !sheet.IsChartSheet).ToArray();

        internal string? CodeName { get; set; }

        internal Dictionary<string, XlsbProjectedPart> RelatedParts { get; } = new(StringComparer.OrdinalIgnoreCase);

        internal Dictionary<string, string> ProjectionPartUris { get; } = new(StringComparer.OrdinalIgnoreCase);

        internal string? VbaProjectPartName { get; set; }

        internal IReadOnlyList<XlsbDefinedName> DefinedNames => _definedNames;

        internal IReadOnlyList<XlsbExternalSheetReference> ExternalSheetReferences => _externalSheetReferences;

        internal IReadOnlyList<bool> SelfSupportingLinks => _selfSupportingLinks;

        internal IReadOnlyList<XlsbImportDiagnostic> Diagnostics => _diagnostics;

        internal IReadOnlyList<XlsbPreservedRecordInfo> PreservedRecords => _preservedRecords;

        internal int SharedStringCount { get; set; }

        internal bool Uses1904DateSystem { get; set; }

        internal XlsbCalculationProperties? CalculationProperties { get; set; }

        internal XlsbWorkbookProtection? WorkbookProtection { get; set; }

        internal XlsbStylesheet? Stylesheet { get; set; }

        internal string? StylesheetPartName { get; set; }

        internal void AddWorksheet(XlsbSheet worksheet) => _worksheets.Add(worksheet);

        internal void AddDefinedName(XlsbDefinedName definedName) {
            _definedNames.Add(definedName);
            _localNamesBySheet = null;
            _allLocalNames = null;
        }

        internal bool HasLocalNameShadow(string name, int sheetIndex) {
            if (_localNamesBySheet == null) {
                var bySheet = new Dictionary<uint, HashSet<string>>();
                var all = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                foreach (XlsbDefinedName candidate in _definedNames) {
                    if (candidate.LocalSheetIndex == uint.MaxValue) continue;
                    all.Add(candidate.Name);
                    if (!bySheet.TryGetValue(candidate.LocalSheetIndex, out HashSet<string>? names)) {
                        names = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        bySheet.Add(candidate.LocalSheetIndex, names);
                    }
                    names.Add(candidate.Name);
                }
                _allLocalNames = all;
                _localNamesBySheet = bySheet;
            }
            return sheetIndex < 0 ? _allLocalNames!.Contains(name)
                : _localNamesBySheet.TryGetValue((uint)sheetIndex, out HashSet<string>? sheetNames) && sheetNames.Contains(name);
        }

        internal void AddExternalSheetReference(XlsbExternalSheetReference reference) => _externalSheetReferences.Add(reference);

        internal void AddSupportingLink(bool isSelf) => _selfSupportingLinks.Add(isSelf);

        internal void AddDiagnostic(XlsbImportDiagnostic diagnostic) => _diagnostics.Add(diagnostic);

        internal void AddPreservedRecord(XlsbPreservedRecordInfo record) => _preservedRecords.Add(record);
    }
}
