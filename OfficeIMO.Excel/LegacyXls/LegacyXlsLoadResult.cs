using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel.LegacyXls {
    /// <summary>
    /// Contains the projected OfficeIMO document and the legacy XLS import report produced from the same parse.
    /// </summary>
    public sealed class LegacyXlsLoadResult : IDisposable {
        private readonly ExcelDocument? _document;

        internal LegacyXlsLoadResult(ExcelDocument? document, LegacyXlsWorkbook workbook, Exception? projectionException = null) {
            _document = document;
            Workbook = workbook ?? throw new ArgumentNullException(nameof(workbook));
            ProjectionException = projectionException;
        }

        /// <summary>
        /// Gets the normal OfficeIMO Excel document projected from supported legacy XLS content.
        /// </summary>
        public ExcelDocument Document => _document ?? throw new InvalidOperationException("No OfficeIMO Excel document was projected from the legacy XLS workbook. Inspect Workbook, Diagnostics, ImportReport, and ProjectionException for import details.", ProjectionException);

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
        public LegacyXlsWorkbook Workbook { get; }

        /// <summary>
        /// Gets diagnostics produced while reading the legacy workbook.
        /// </summary>
        public IReadOnlyList<LegacyXlsImportDiagnostic> Diagnostics => Workbook.Diagnostics;

        /// <summary>
        /// Gets unsupported or preserve-only features discovered during import.
        /// </summary>
        public IReadOnlyList<LegacyXlsUnsupportedFeature> UnsupportedFeatures => Workbook.UnsupportedFeatures;

        /// <summary>
        /// Gets a compact import report for corpus baselines and preflight checks.
        /// </summary>
        public LegacyXlsImportReport ImportReport => Workbook.CreateImportReport();

        /// <summary>
        /// Gets whether the legacy XLS import produced error diagnostics.
        /// </summary>
        public bool HasImportErrors => Diagnostics.Any(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error);

        /// <summary>
        /// Gets whether the legacy XLS import discovered unsupported or preserve-only features.
        /// </summary>
        public bool HasUnsupportedFeatures => UnsupportedFeatures.Count > 0;

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
                throw new InvalidOperationException("Legacy XLS import discovered unsupported or preserve-only features: " + FormatUnsupportedFeatures(UnsupportedFeatures));
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

        private static string FormatUnsupportedFeatures(IEnumerable<LegacyXlsUnsupportedFeature> features) {
            return string.Join("; ", features.Take(8).Select(feature => {
                string sheet = feature.SheetName == null ? string.Empty : $" [{feature.SheetName}]";
                string record = feature.RecordType == null ? string.Empty : $" record=0x{feature.RecordType.Value:X4}";
                string offset = feature.RecordOffset == null ? string.Empty : $" offset={feature.RecordOffset.Value}";
                return $"{feature.Code}{sheet}{record}{offset}: {feature.Description}";
            }));
        }
    }
}
