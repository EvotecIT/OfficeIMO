using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls {
    /// <summary>
    /// Contains the projected OfficeIMO document and the legacy XLS import report produced from the same parse.
    /// </summary>
    public sealed class LegacyXlsLoadResult : IDisposable {
        private readonly ExcelDocument? _document;
        private readonly Lazy<LegacyXlsImportReport> _importReport;
        private readonly Lazy<LegacyXlsImportSummary> _summary;

        internal LegacyXlsLoadResult(ExcelDocument? document, LegacyXlsWorkbook workbook, Exception? projectionException = null) {
            _document = document;
            Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            ProjectionException = projectionException;
            _importReport = new Lazy<LegacyXlsImportReport>(() => Workbook.CreateImportReport());
            _summary = new Lazy<LegacyXlsImportSummary>(() => new LegacyXlsImportSummary(this));
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
        public bool HasImportErrors => Diagnostics.Any(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);

        /// <summary>
        /// Gets whether the legacy XLS import discovered unsupported or preserve-only features.
        /// </summary>
        public bool HasUnsupportedFeatures => UnsupportedFeatures.Count > 0 || PreservedFeatures.Count > 0;

        /// <summary>Gets whether conversion to XLSX would omit known legacy content.</summary>
        public bool HasConversionLoss => UnsupportedFeatures.Count > 0
            || PreservedFeatures.Count > 0
            || UnsupportedSheets.Count > 0
            || CompoundFeatures.Any(feature =>
                feature.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject
                || feature.Kind == LegacyXlsCompoundFeatureRecordKind.OleObject);

        /// <summary>
        /// Throws when the legacy XLS import produced error diagnostics.
        /// </summary>
        public LegacyXlsLoadResult EnsureNoImportErrors() {
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
    }
}
