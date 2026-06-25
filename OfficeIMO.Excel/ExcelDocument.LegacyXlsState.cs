using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private LegacyXlsImportDiagnostic[] _legacyXlsImportDiagnostics = Array.Empty<LegacyXlsImportDiagnostic>();
        private LegacyXlsUnsupportedFeature[] _legacyXlsUnsupportedFeatures = Array.Empty<LegacyXlsUnsupportedFeature>();
        private LegacyXlsUnsupportedSheet[] _legacyXlsUnsupportedSheets = Array.Empty<LegacyXlsUnsupportedSheet>();
        private string? _legacyXlsSourcePath;

        /// <summary>
        /// Gets whether this workbook was projected from a legacy binary XLS source through normal loading.
        /// </summary>
        public bool WasLoadedFromLegacyXls { get; private set; }

        /// <summary>
        /// Gets diagnostics produced while importing a legacy binary XLS source through normal loading.
        /// </summary>
        public IReadOnlyList<LegacyXlsImportDiagnostic> LegacyXlsImportDiagnostics => _legacyXlsImportDiagnostics;

        /// <summary>
        /// Gets unsupported or preserve-only legacy XLS features discovered during normal loading.
        /// </summary>
        public IReadOnlyList<LegacyXlsUnsupportedFeature> LegacyXlsUnsupportedFeatures => _legacyXlsUnsupportedFeatures;

        /// <summary>
        /// Gets legacy XLS sheet entries that were discovered but not projected as normal worksheets.
        /// </summary>
        public IReadOnlyList<LegacyXlsUnsupportedSheet> LegacyXlsUnsupportedSheets => _legacyXlsUnsupportedSheets;

        internal void MarkLoadedFromLegacyXls(string? sourcePath, LegacyXlsWorkbook workbook) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            WasLoadedFromLegacyXls = true;
            _legacyXlsSourcePath = sourcePath;
            _legacyXlsImportDiagnostics = workbook.Diagnostics.ToArray();
            _legacyXlsUnsupportedFeatures = workbook.UnsupportedFeatures.ToArray();
            _legacyXlsUnsupportedSheets = workbook.UnsupportedSheets.ToArray();

            if (!string.IsNullOrEmpty(sourcePath)) {
                FilePath = sourcePath!;
            }
        }

        private void EnsureLegacyXlsCanSaveToPath(string path) {
            if (!WasLoadedFromLegacyXls) {
                return;
            }

            if (!ExcelDocumentLoadRouting.HasLegacyXlsExtension(path)) {
                return;
            }

            string source = string.IsNullOrWhiteSpace(_legacyXlsSourcePath)
                ? "a legacy binary .xls source"
                : $"legacy binary .xls source '{_legacyXlsSourcePath}'";
            throw new NotSupportedException($"Native XLS saving is not supported. This workbook was loaded from {source}; save it to an .xlsx path instead.");
        }
    }
}
