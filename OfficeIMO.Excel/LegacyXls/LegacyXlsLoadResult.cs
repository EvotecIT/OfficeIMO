using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls {
    /// <summary>
    /// Contains the projected OfficeIMO document and the legacy XLS import report produced from the same parse.
    /// </summary>
    public sealed class LegacyXlsLoadResult : IDisposable, IOfficeConversionReport {
        private readonly ExcelDocument? _document;
        private readonly Lazy<LegacyXlsImportReport> _importReport;
        private readonly Lazy<LegacyXlsImportSummary> _summary;
        private readonly Lazy<IReadOnlyList<OfficeConversionFidelityDiagnostic>> _fidelityDiagnostics;

        internal LegacyXlsLoadResult(ExcelDocument? document, LegacyXlsWorkbook workbook, Exception? projectionException = null) {
            _document = document;
            Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            ProjectionException = projectionException;
            _importReport = new Lazy<LegacyXlsImportReport>(() => Workbook.CreateImportReport());
            _summary = new Lazy<LegacyXlsImportSummary>(() => new LegacyXlsImportSummary(this));
            _fidelityDiagnostics = new Lazy<IReadOnlyList<OfficeConversionFidelityDiagnostic>>(CreateFidelityDiagnostics);
        }

        /// <summary>
        /// Gets the normal OfficeIMO Excel document projected from supported legacy XLS content.
        /// </summary>
        public ExcelDocument Document => _document ?? throw new InvalidOperationException("No OfficeIMO Excel document was projected from the legacy XLS workbook. Inspect AdvancedWorkbook, Diagnostics, CreateImportReport(), and ProjectionException for import details.", ProjectionException);

        /// <summary>
        /// Gets whether supported legacy XLS content was projected into a normal OfficeIMO Excel document.
        /// </summary>
        public bool HasDocument => _document != null;

        /// <summary>
        /// Gets the projection failure captured while preserving parser diagnostics for report callers.
        /// </summary>
        public Exception? ProjectionException { get; }

        /// <summary>
        /// Gets the neutral legacy XLS workbook model produced by the parser.
        /// </summary>
        internal LegacyXlsWorkbook Workbook { get; }

        /// <summary>Gets the advanced neutral parser model for forensic or corpus analysis.</summary>
        public LegacyXlsWorkbook AdvancedWorkbook => Workbook;

        /// <summary>
        /// Gets diagnostics produced while reading the legacy workbook.
        /// </summary>
        public IReadOnlyList<LegacyXlsImportDiagnostic> Diagnostics => Workbook.Diagnostics;

        /// <summary>
        /// Gets unsupported or preserve-only features discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsUnsupportedFeature> UnsupportedFeatures => Workbook.UnsupportedFeatures;

        /// <summary>Gets preserve-only BIFF feature records that were not projected into the normal workbook model.</summary>
        public IReadOnlyList<LegacyXlsPreservedFeatureRecord> PreservedFeatures => Workbook.PreservedFeatureRecords;

        /// <summary>Gets sheet entries that were not projected as normal worksheets.</summary>
        public IReadOnlyList<LegacyXlsUnsupportedSheet> UnsupportedSheets => Workbook.UnsupportedSheets;

        /// <summary>Gets chart sheets that were projected into chart-sheet package parts.</summary>
        public IReadOnlyList<LegacyXlsChartSheet> ChartSheets => Workbook.ChartSheets;

        /// <summary>Gets preserve-only features found in the OLE compound container.</summary>
        public IReadOnlyList<LegacyXlsCompoundFeatureRecord> CompoundFeatures => Workbook.CompoundFeatureRecords;

        /// <summary>
        /// Gets the corpus-grade import report used by OfficeIMO's compatibility tests.
        /// </summary>
        internal LegacyXlsImportReport ImportReport => _importReport.Value;

        /// <summary>Gets a compact cached summary intended for normal application code.</summary>
        public LegacyXlsImportSummary Summary => _summary.Value;

        /// <summary>Creates or returns a compact cached import report for preflight checks and diagnostics.</summary>
        public LegacyXlsImportReport CreateImportReport() => _importReport.Value;

        /// <summary>
        /// Gets whether the legacy XLS import produced error diagnostics.
        /// </summary>
        public bool HasImportErrors => ProjectionException != null ||
            Diagnostics.Any(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);

        /// <summary>
        /// Gets whether the legacy XLS import discovered unsupported or preserve-only features.
        /// </summary>
        public bool HasUnsupportedFeatures => UnsupportedFeatures.Count > 0 || PreservedFeatures.Count > 0;

        /// <summary>Gets whether conversion to XLSX would omit known legacy content.</summary>
        public bool HasConversionLoss => HasLoss;

        /// <inheritdoc />
        public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _fidelityDiagnostics.Value;

        /// <inheritdoc />
        public bool HasLoss => FidelityDiagnostics.Any(diagnostic =>
            diagnostic.LossKind != OfficeConversionLossKind.None);

        /// <inheritdoc />
        public void RequireNoLoss() {
            if (HasLoss) throw new InvalidDataException(
                "The legacy XLS import reported content loss. Inspect FidelityDiagnostics and the source import collections for details.");
        }

        /// <summary>
        /// Throws when the legacy XLS import produced error diagnostics.
        /// </summary>
        public LegacyXlsLoadResult EnsureNoImportErrors() {
            if (ProjectionException != null) {
                throw new InvalidOperationException(
                    "Legacy XLS content was parsed but could not be projected to an OfficeIMO workbook.",
                    ProjectionException);
            }
            if (HasImportErrors) {
                throw new InvalidOperationException("Legacy XLS import produced errors: " + FormatDiagnostics(Diagnostics.Where(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error)));
            }

            return this;
        }

        /// <summary>
        /// Throws when the legacy XLS import discovered unsupported or preserve-only features.
        /// </summary>
        public LegacyXlsLoadResult EnsureNoUnsupportedFeatures() {
            if (HasUnsupportedFeatures) {
                throw new InvalidOperationException(
                    "Legacy XLS import discovered unsupported or preserve-only features: "
                    + FormatUnsupportedFeatures(UnsupportedFeatures, PreservedFeatures));
            }

            return this;
        }

        /// <summary>Throws when conversion to XLSX would omit known legacy content.</summary>
        public LegacyXlsLoadResult EnsureNoConversionLoss() {
            if (HasConversionLoss) {
                throw new InvalidOperationException("Legacy XLS import contains unsupported sheets, unsupported or preserve-only features, VBA, or OLE content that cannot be projected to XLSX without loss.");
            }

            return this;
        }

        /// <summary>
        /// Disposes the projected OfficeIMO document.
        /// </summary>
        public void Dispose() {
            _document?.Dispose();
        }

        private static string FormatDiagnostics(IEnumerable<LegacyXlsImportDiagnostic> diagnostics) {
            return string.Join("; ", diagnostics.Take(8).Select(diagnostic => diagnostic.ToString()));
        }

        private static string FormatUnsupportedFeatures(
            IEnumerable<LegacyXlsUnsupportedFeature> features,
            IEnumerable<LegacyXlsPreservedFeatureRecord> preservedFeatures) {
            IEnumerable<string> unsupported = features.Select(feature => {
                string sheet = feature.SheetName == null ? string.Empty : $" [{feature.SheetName}]";
                string record = feature.RecordType == null ? string.Empty : $" record=0x{feature.RecordType.Value:X4}";
                string offset = feature.RecordOffset == null ? string.Empty : $" offset={feature.RecordOffset.Value}";
                return $"{feature.Code}{sheet}{record}{offset}: {feature.Description}";
            });
            IEnumerable<string> preserved = preservedFeatures.Select(feature => {
                string sheet = feature.SheetName == null ? string.Empty : $" [{feature.SheetName}]";
                return $"{feature.Code}{sheet} record=0x{feature.RecordType:X4} offset={feature.RecordOffset}: {feature.Description}";
            });
            return string.Join("; ", unsupported.Concat(preserved).Distinct(StringComparer.Ordinal).Take(8));
        }

        private IReadOnlyList<OfficeConversionFidelityDiagnostic> CreateFidelityDiagnostics() {
            var diagnostics = new List<OfficeConversionFidelityDiagnostic>();
            if (ProjectionException != null) {
                diagnostics.Add(new OfficeConversionFidelityDiagnostic(
                    "XLS-PROJECTION-FAILED",
                    ProjectionException.Message,
                    OfficeConversionLossKind.Failure,
                    "OfficeIMO.Excel.LegacyXls.Projection",
                    ProjectionException.GetType().FullName));
            }
            diagnostics.AddRange(Diagnostics.Select(diagnostic => new OfficeConversionFidelityDiagnostic(
                diagnostic.Code,
                diagnostic.Message,
                ClassifyLoss(diagnostic),
                "OfficeIMO.Excel.LegacyXls.Reader",
                FormatLocation(diagnostic.SheetName, diagnostic.RecordOffset))));
            diagnostics.AddRange(UnsupportedFeatures.Select(feature => new OfficeConversionFidelityDiagnostic(
                feature.Code, feature.Description, OfficeConversionLossKind.Omission,
                "OfficeIMO.Excel.LegacyXls.Reader", FormatLocation(feature.SheetName, feature.RecordOffset))));
            diagnostics.AddRange(PreservedFeatures.Select(feature => new OfficeConversionFidelityDiagnostic(
                feature.Code, feature.Description, OfficeConversionLossKind.Omission,
                "OfficeIMO.Excel.LegacyXls.Reader", FormatLocation(feature.SheetName, feature.RecordOffset))));
            diagnostics.AddRange(UnsupportedSheets.Select(sheet => new OfficeConversionFidelityDiagnostic(
                "XLS-UNSUPPORTED-SHEET", $"Sheet '{sheet.Name}' is not projected as an editable worksheet.",
                OfficeConversionLossKind.Omission, "OfficeIMO.Excel.LegacyXls.Reader",
                FormatLocation(sheet.Name, sheet.StreamOffset))));
            diagnostics.AddRange(CompoundFeatures
                .Where(feature => feature.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject
                    || feature.Kind == LegacyXlsCompoundFeatureRecordKind.OleObject)
                .Select(feature => new OfficeConversionFidelityDiagnostic(
                    "XLS-COMPOUND-" + feature.Kind.ToString().ToUpperInvariant(),
                    $"The {feature.Kind} compound feature is preserved but not projected to XLSX.",
                    OfficeConversionLossKind.Omission,
                    "OfficeIMO.Excel.LegacyXls.Reader",
                    feature.Entries.FirstOrDefault())));
            return Array.AsReadOnly(diagnostics.ToArray());
        }

        private static OfficeConversionLossKind ClassifyLoss(LegacyXlsImportDiagnostic diagnostic) {
            if (diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error) return OfficeConversionLossKind.Failure;
            // An unread shared string can become an empty LabelSst cell; it is not a
            // representational approximation of the original value.
            if (diagnostic.Code is "XLS-BIFF-SST-SHORT" or "XLS-BIFF-SST-STRING-INVALID")
                return OfficeConversionLossKind.Omission;
            return diagnostic.Severity == LegacyXlsDiagnosticSeverity.Warning
                ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.None;
        }

        private static string? FormatLocation(string? sheetName, int? recordOffset) {
            if (!string.IsNullOrWhiteSpace(sheetName) && recordOffset.HasValue) {
                return $"sheet:{sheetName}/offset:{recordOffset.Value}";
            }
            if (!string.IsNullOrWhiteSpace(sheetName)) return "sheet:" + sheetName;
            return recordOffset.HasValue ? "offset:" + recordOffset.Value : null;
        }
    }
}
