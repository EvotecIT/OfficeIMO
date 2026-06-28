using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Threading;
using System.Threading.Tasks;

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

        internal static ExcelDocument ProjectLoadedLegacyXlsWorkbook(LegacyXlsWorkbook workbook, string? sourcePath) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            if (workbook.Worksheets.Count == 0) {
                throw new InvalidDataException("Legacy XLS import failed: no supported worksheets were projected. Unsupported legacy sheet content cannot be saved as a normal .xlsx workbook.");
            }

            ExcelDocument document = workbook.ToExcelDocument();
            document.MarkLoadedFromLegacyXls(sourcePath, workbook);
            return document;
        }

        private void EnsureLegacyBinaryExcelSaveTargetSupported(string path, bool allowNativeXls) {
            if (ExcelDocumentLoadRouting.HasLegacyXlsExtension(path)) {
                if (allowNativeXls) {
                    return;
                }

                throw new NotSupportedException("Native XLS encrypted saving is not supported. Save encrypted workbooks to an .xlsx path instead.");
            }

            if (!ExcelDocumentLoadRouting.HasLegacyBinaryExcelExtension(path)) {
                return;
            }

            if (!WasLoadedFromLegacyXls) {
                throw new NotSupportedException("Native XLS saving currently supports .xls workbook files only. Legacy .xlt, .xla, .xlm, and .xlw save targets are not supported.");
            }

            string source = string.IsNullOrWhiteSpace(_legacyXlsSourcePath)
                ? "a legacy binary Excel source"
                : $"legacy binary Excel source '{_legacyXlsSourcePath}'";
            throw new NotSupportedException($"Native XLS saving currently supports .xls workbook files only. This workbook was loaded from {source}; legacy .xlt, .xla, .xlm, and .xlw save targets are not supported.");
        }

        private bool TrySaveNativeLegacyXlsToFile(string path, bool openExcel, ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            if (!ExcelDocumentLoadRouting.HasLegacyXlsExtension(path)) {
                return false;
            }

            cancellationToken.ThrowIfCancellationRequested();
            PrepareWorkbookForSave(options);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] xlsBytes = OfficeIMO.Excel.LegacyXls.Write.LegacyXlsWriter.WriteWorkbook(this);
            cancellationToken.ThrowIfCancellationRequested();
            CommitPreparedPackageToFile(path, xlsBytes);
            FilePath = path;
            MarkPackageClean(null);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Native XLS save used the first-party BIFF8 writer.");

            if (openExcel) {
                Helpers.Open(path, true);
            }

            return true;
        }

        private async Task<bool> TrySaveNativeLegacyXlsToFileAsync(string path, bool openExcel, ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            if (!ExcelDocumentLoadRouting.HasLegacyXlsExtension(path)) {
                return false;
            }

            cancellationToken.ThrowIfCancellationRequested();
            PrepareWorkbookForSave(options);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] xlsBytes = OfficeIMO.Excel.LegacyXls.Write.LegacyXlsWriter.WriteWorkbook(this);
            cancellationToken.ThrowIfCancellationRequested();
            await CommitPreparedPackageToFileAsync(path, xlsBytes, cancellationToken).ConfigureAwait(false);
            FilePath = path;
            MarkPackageClean(null);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Native XLS save used the first-party BIFF8 writer.");

            if (openExcel) {
                Open(path, true);
            }

            return true;
        }

        private bool TrySaveNativeLegacyXlsToStream(Stream destination, ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            if (options?.StreamFormat != ExcelStreamSaveFormat.LegacyXls) {
                return false;
            }

            cancellationToken.ThrowIfCancellationRequested();
            PrepareWorkbookForSave(options);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] xlsBytes = OfficeIMO.Excel.LegacyXls.Write.LegacyXlsWriter.WriteWorkbook(this);
            cancellationToken.ThrowIfCancellationRequested();
            PrepareDestinationStreamForWrite(destination);
            destination.Write(xlsBytes, 0, xlsBytes.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            MarkPackageClean(null);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Native XLS stream save used the first-party BIFF8 writer.");
            return true;
        }

        private async Task<bool> TrySaveNativeLegacyXlsToStreamAsync(Stream destination, ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            if (options?.StreamFormat != ExcelStreamSaveFormat.LegacyXls) {
                return false;
            }

            cancellationToken.ThrowIfCancellationRequested();
            PrepareWorkbookForSave(options);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] xlsBytes = OfficeIMO.Excel.LegacyXls.Write.LegacyXlsWriter.WriteWorkbook(this);
            cancellationToken.ThrowIfCancellationRequested();
            PrepareDestinationStreamForWrite(destination);
            await destination.WriteAsync(xlsBytes, 0, xlsBytes.Length, cancellationToken).ConfigureAwait(false);
            try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
            MarkPackageClean(null);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Native XLS stream save used the first-party BIFF8 writer.");
            return true;
        }

        private static void EnsureLegacyBinaryEncryptedSaveTargetSupported(string path) {
            if (!ExcelDocumentLoadRouting.HasLegacyBinaryExcelExtension(path)) {
                return;
            }

            throw new NotSupportedException("Encrypted saving is supported for Office Open XML workbooks only. Save encrypted workbooks to an .xlsx path instead.");
        }
    }
}
