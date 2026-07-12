using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using System.Threading;
using System.Threading.Tasks;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument {
        private LegacyXlsImportDiagnostic[] _legacyXlsImportDiagnostics = Array.Empty<LegacyXlsImportDiagnostic>();
        private LegacyXlsUnsupportedFeature[] _legacyXlsUnsupportedFeatures = Array.Empty<LegacyXlsUnsupportedFeature>();
        private LegacyXlsPreservedFeatureRecord[] _legacyXlsPreservedFeatures = Array.Empty<LegacyXlsPreservedFeatureRecord>();
        private LegacyXlsUnsupportedSheet[] _legacyXlsUnsupportedSheets = Array.Empty<LegacyXlsUnsupportedSheet>();
        private LegacyXlsChartSheet[] _legacyXlsChartSheets = Array.Empty<LegacyXlsChartSheet>();
        private LegacyXlsCompoundFeatureRecord[] _legacyXlsCompoundFeatures = Array.Empty<LegacyXlsCompoundFeatureRecord>();
        private string? _legacyXlsSourcePath;

        /// <summary>Gets the detected physical format of the workbook source.</summary>
        public ExcelFileFormat SourceFormat { get; private set; } = ExcelFileFormat.Xlsx;

        /// <summary>Gets the original legacy source path, or the current Open XML file association.</summary>
        public string? SourcePath => SourceFormat == ExcelFileFormat.Xls
            ? _legacyXlsSourcePath
            : string.IsNullOrWhiteSpace(FilePath) ? null : FilePath;

        /// <summary>
        /// Gets diagnostics produced while importing a legacy binary XLS source through normal loading.
        /// </summary>
        public IReadOnlyList<LegacyXlsImportDiagnostic> LegacyXlsImportDiagnostics => _legacyXlsImportDiagnostics;

        /// <summary>
        /// Gets unsupported or preserve-only legacy XLS features discovered during normal loading.
        /// </summary>
        public IReadOnlyList<LegacyXlsUnsupportedFeature> LegacyXlsUnsupportedFeatures => _legacyXlsUnsupportedFeatures;

        /// <summary>Gets preserve-only BIFF feature records that were not projected into the normal workbook model.</summary>
        public IReadOnlyList<LegacyXlsPreservedFeatureRecord> LegacyXlsPreservedFeatures => _legacyXlsPreservedFeatures;

        /// <summary>
        /// Gets legacy XLS sheet entries that were discovered but not projected as normal worksheets.
        /// </summary>
        public IReadOnlyList<LegacyXlsUnsupportedSheet> LegacyXlsUnsupportedSheets => _legacyXlsUnsupportedSheets;

        /// <summary>
        /// Gets legacy XLS chart sheets decoded during import and projected as chart-sheet package parts.
        /// </summary>
        public IReadOnlyList<LegacyXlsChartSheet> LegacyXlsChartSheets => _legacyXlsChartSheets;

        /// <summary>
        /// Gets legacy XLS compound-container features decoded during import but not projected into the normal workbook package.
        /// </summary>
        public IReadOnlyList<LegacyXlsCompoundFeatureRecord> LegacyXlsCompoundFeatures => _legacyXlsCompoundFeatures;

        internal void MarkLoadedFromLegacyXls(string? sourcePath, LegacyXlsWorkbook workbook) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            SourceFormat = ExcelFileFormat.Xls;
            _legacyXlsSourcePath = sourcePath;
            _legacyXlsImportDiagnostics = workbook.Diagnostics.ToArray();
            _legacyXlsUnsupportedFeatures = workbook.UnsupportedFeatures.ToArray();
            _legacyXlsPreservedFeatures = workbook.PreservedFeatureRecords.ToArray();
            _legacyXlsUnsupportedSheets = workbook.UnsupportedSheets.ToArray();
            _legacyXlsChartSheets = workbook.ChartSheets.ToArray();
            _legacyXlsCompoundFeatures = workbook.CompoundFeatureRecords.ToArray();

            if (!string.IsNullOrEmpty(sourcePath)) {
                FilePath = sourcePath!;
            }
        }

        internal static ExcelDocument ProjectLoadedLegacyXlsWorkbook(LegacyXlsWorkbook workbook, string? sourcePath) {
            if (workbook == null) throw new ArgumentNullException(nameof(workbook));

            if (workbook.Worksheets.Count == 0 && workbook.ChartSheets.Count == 0) {
                throw new InvalidDataException("Legacy XLS import failed: no supported worksheets or chart sheets were projected. Unsupported legacy sheet content cannot be saved as a normal .xlsx workbook.");
            }

            ExcelDocument document = workbook.ToExcelDocument();
            document.MarkLoadedFromLegacyXls(sourcePath, workbook);
            return document;
        }

        private void EnsureLegacyBinaryExcelSaveTargetSupported(string path, bool allowNativeXls, ExcelSaveOptions? options = null) {
            if (ExcelDocumentLoadRouting.HasLegacyXlsExtension(path)) {
                EnsureLegacyXlsSaveDoesNotDropImportedContent(options, includeProjectedChartSheets: true);

                if (allowNativeXls) {
                    return;
                }

                throw new NotSupportedException("Native XLS encrypted saving is not supported. Save encrypted workbooks to an .xlsx path instead.");
            }

            if (!ExcelDocumentLoadRouting.HasLegacyBinaryExcelExtension(path)) {
                return;
            }

            if (SourceFormat != ExcelFileFormat.Xls) {
                throw new NotSupportedException("Native XLS saving currently supports .xls workbook files only. Legacy .xlt, .xla, .xlm, and .xlw save targets are not supported.");
            }

            string source = string.IsNullOrWhiteSpace(_legacyXlsSourcePath)
                ? "a legacy binary Excel source"
                : $"legacy binary Excel source '{_legacyXlsSourcePath}'";
            throw new NotSupportedException($"Native XLS saving currently supports .xls workbook files only. This workbook was loaded from {source}; legacy .xlt, .xla, .xlm, and .xlw save targets are not supported.");
        }

        private bool HasLossyLegacyXlsImportState(bool includeProjectedChartSheets) {
            return _legacyXlsUnsupportedFeatures.Length > 0
                || _legacyXlsPreservedFeatures.Length > 0
                || _legacyXlsUnsupportedSheets.Length > 0
                || (includeProjectedChartSheets && _legacyXlsChartSheets.Length > 0)
                || _legacyXlsCompoundFeatures.Any(feature =>
                    feature.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject
                    || feature.Kind == LegacyXlsCompoundFeatureRecordKind.OleObject);
        }

        private void EnsureLegacyXlsSaveDoesNotDropImportedContent(ExcelSaveOptions? options, bool includeProjectedChartSheets = false) {
            if (SourceFormat != ExcelFileFormat.Xls
                || !HasLossyLegacyXlsImportState(includeProjectedChartSheets)
                || options?.LossPolicy == ExcelConversionLossPolicy.Allow) {
                return;
            }

            string source = string.IsNullOrWhiteSpace(_legacyXlsSourcePath)
                ? "a legacy binary .xls source"
                : $"legacy binary .xls source '{_legacyXlsSourcePath}'";
            string codes = string.Join(", ", _legacyXlsUnsupportedFeatures
                .Select(feature => feature.Code)
                .Concat(_legacyXlsPreservedFeatures.Select(feature => feature.Code))
                .Concat(_legacyXlsUnsupportedSheets.Select(sheet => "UnsupportedSheet:" + sheet.Kind))
                .Concat(_legacyXlsCompoundFeatures
                    .Where(feature => feature.Kind == LegacyXlsCompoundFeatureRecordKind.VbaProject || feature.Kind == LegacyXlsCompoundFeatureRecordKind.OleObject)
                    .Select(feature => "Compound:" + feature.Kind))
                .Where(code => !string.IsNullOrWhiteSpace(code))
                .Distinct(StringComparer.Ordinal)
                .Take(8));
            string details = string.IsNullOrWhiteSpace(codes) ? string.Empty : $" Findings: {codes}.";
            throw new NotSupportedException($"Saving is blocked because this workbook was loaded from {source} with unsupported, preserve-only, or non-projected legacy content.{details} Review LegacyXlsUnsupportedFeatures, LegacyXlsPreservedFeatures, LegacyXlsUnsupportedSheets, and LegacyXlsCompoundFeatures, or set ExcelSaveOptions.LossPolicy to ExcelConversionLossPolicy.Allow when that loss is intentional.");
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
            bool reopenWorkingPackage = ShouldCloseOpenPackageForNativeLegacyXlsFileSave(path);
            byte[]? workingPackageBytes = reopenWorkingPackage
                ? CaptureOpenXmlPackageBytesForNativeLegacyXlsReopen()
                : null;
            if (reopenWorkingPackage) {
                CloseOpenPackageForNativeLegacyXlsSave();
            }

            try {
                CommitPreparedPackageToFile(path, xlsBytes);
            } catch {
                RestorePackageAfterFailedNativeLegacyXlsFileCommit(workingPackageBytes);
                throw;
            }

            FilePath = path;
            DisablePackageCopyBackAfterNativeLegacyXlsFileSave();
            if (workingPackageBytes != null) {
                ReloadFromBytes(workingPackageBytes);
            } else {
                MarkPackageClean(null);
            }

            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Native XLS save used the first-party BIFF8 writer.");

            if (openExcel) {
                OfficeIMO.Core.OfficeFileLauncher.Open(path);
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
            bool reopenWorkingPackage = ShouldCloseOpenPackageForNativeLegacyXlsFileSave(path);
            byte[]? workingPackageBytes = reopenWorkingPackage
                ? CaptureOpenXmlPackageBytesForNativeLegacyXlsReopen()
                : null;
            if (reopenWorkingPackage) {
                CloseOpenPackageForNativeLegacyXlsSave();
            }

            try {
                await CommitPreparedPackageToFileAsync(path, xlsBytes, cancellationToken).ConfigureAwait(false);
            } catch {
                RestorePackageAfterFailedNativeLegacyXlsFileCommit(workingPackageBytes);
                throw;
            }

            FilePath = path;
            DisablePackageCopyBackAfterNativeLegacyXlsFileSave();
            if (workingPackageBytes != null) {
                ReloadFromBytes(workingPackageBytes);
            } else {
                MarkPackageClean(null);
            }

            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Native XLS save used the first-party BIFF8 writer.");

            if (openExcel) {
                OpenInApplication(path);
            }

            return true;
        }

        private bool TrySaveNativeLegacyXlsToStream(Stream destination, ExcelFileFormat format, ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            if (format != ExcelFileFormat.Xls) {
                return false;
            }

            EnsureLegacyXlsSaveDoesNotDropImportedContent(options, includeProjectedChartSheets: true);
            cancellationToken.ThrowIfCancellationRequested();
            PrepareWorkbookForSave(options);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] xlsBytes = OfficeIMO.Excel.LegacyXls.Write.LegacyXlsWriter.WriteWorkbook(this);
            cancellationToken.ThrowIfCancellationRequested();
            PrepareDestinationStreamForWrite(destination);
            destination.Write(xlsBytes, 0, xlsBytes.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }
            MarkPackageClean(null);
            DisablePackageCopyBackAfterNativeLegacyXlsSave(destination);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Native XLS stream save used the first-party BIFF8 writer.");
            return true;
        }

        private async Task<bool> TrySaveNativeLegacyXlsToStreamAsync(Stream destination, ExcelFileFormat format, ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            if (format != ExcelFileFormat.Xls) {
                return false;
            }

            EnsureLegacyXlsSaveDoesNotDropImportedContent(options, includeProjectedChartSheets: true);
            cancellationToken.ThrowIfCancellationRequested();
            PrepareWorkbookForSave(options);
            cancellationToken.ThrowIfCancellationRequested();
            byte[] xlsBytes = OfficeIMO.Excel.LegacyXls.Write.LegacyXlsWriter.WriteWorkbook(this);
            cancellationToken.ThrowIfCancellationRequested();
            PrepareDestinationStreamForWrite(destination);
            await destination.WriteAsync(xlsBytes, 0, xlsBytes.Length, cancellationToken).ConfigureAwait(false);
            try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
            MarkPackageClean(null);
            DisablePackageCopyBackAfterNativeLegacyXlsSave(destination);
            LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Native XLS stream save used the first-party BIFF8 writer.");
            return true;
        }

        private byte[] CaptureOpenXmlPackageBytesForNativeLegacyXlsReopen() {
            using var snapshot = new MemoryStream();
            using (_spreadSheetDocument.Clone(snapshot)) { }
            return snapshot.ToArray();
        }

        private void RestorePackageAfterFailedNativeLegacyXlsFileCommit(byte[]? workingPackageBytes) {
            if (workingPackageBytes == null) {
                return;
            }

            try {
                ReloadFromBytes(workingPackageBytes);
            } catch {
            }
        }

        private bool ShouldCloseOpenPackageForNativeLegacyXlsFileSave(string path) {
            if (SourceFormat == ExcelFileFormat.Xls || string.IsNullOrWhiteSpace(FilePath)) {
                return false;
            }

            return true;
        }

        private void CloseOpenPackageForNativeLegacyXlsSave() {
            CloseSpreadsheetDocumentAfterNativeLegacyXlsSave();
            DisposePackageStreamAfterNativeLegacyXlsSave(disposeSourceStream: false);
        }

        private void DisablePackageCopyBackAfterNativeLegacyXlsFileSave() {
            _sourceStream = null;
            _copyPackageToSourceOnDispose = false;
            _copyPackageToFilePathOnDispose = false;
            _leaveSourceStreamOpen = true;
        }

        private void DisablePackageCopyBackAfterNativeLegacyXlsSave(Stream destination) {
            if (!ReferenceEquals(destination, _sourceStream)) {
                return;
            }

            _sourceStream = null;
            _copyPackageToSourceOnDispose = false;
            _copyPackageToFilePathOnDispose = false;
            _leaveSourceStreamOpen = true;
        }

        private void CloseSpreadsheetDocumentAfterNativeLegacyXlsSave() {
            if (_spreadSheetDocument != null) {
                try { _spreadSheetDocument.Dispose(); } catch { }
                _spreadSheetDocument = null!;
            }
        }

        private void DisposePackageStreamAfterNativeLegacyXlsSave(bool disposeSourceStream) {
            if (_packageStream != null) {
                DisposeStream(_packageStream);
                _packageStream = null;
            }

            if (disposeSourceStream && _sourceStream != null) {
                try { _sourceStream.Dispose(); } catch { }
            }

            _sourceStream = null;
            _copyPackageToSourceOnDispose = false;
            _copyPackageToFilePathOnDispose = false;
            _leaveSourceStreamOpen = true;
        }

        private static void EnsureLegacyBinaryEncryptedSaveTargetSupported(string path) {
            if (!ExcelDocumentLoadRouting.HasLegacyBinaryExcelExtension(path)) {
                return;
            }

            throw new NotSupportedException("Encrypted saving is supported for Office Open XML workbooks only. Save encrypted workbooks to an .xlsx path instead.");
        }
    }
}
