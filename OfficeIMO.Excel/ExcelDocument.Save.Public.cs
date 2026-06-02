using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Shared;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        /// <summary>
        /// Opens the document with the associated application.
        /// </summary>
        /// <param name="filePath">Optional path to open.</param>
        /// <param name="openExcel">Whether to launch Excel.</param>
        public void Open(string filePath = "", bool openExcel = true) {
            if (filePath == "") {
                filePath = this.FilePath;
            }
            Helpers.Open(filePath, openExcel);
        }

        /// <summary>
        /// Performs a safety preflight across all worksheets to reduce the likelihood of Excel prompting
        /// for repairs on open. It removes empty containers (Hyperlinks/MergeCells), drops orphaned drawing
        /// and header/footer references, and cleans up invalid table references.
        /// </summary>
        public void PreflightWorkbook() {
            MaterializePendingDirectCellValueSheetIfNeeded();
            PreflightWorkbook(Sheets);
            _requiresSavePreflight = false;
        }

        private void PreflightWorkbook(IEnumerable<ExcelSheet> sheets) {
            foreach (var sheet in sheets) {
                sheet.Preflight();
            }
            CleanupWorkbookViewArtifacts(save: true);
            CleanupStyleAndSharedStringArtifacts(save: true);
            CleanupCalculationArtifacts(save: true);
            CleanupDefinedNameArtifacts(includeAggressiveRepairs: false, save: true);
        }

        /// <summary>
        /// Closes the underlying spreadsheet document.
        /// </summary>
        public void Close() {
            Dispose();
        }

        private static void EnsureDirectoryWritable(string path) {
            if (string.IsNullOrWhiteSpace(path)) {
                return;
            }

            var directory = Path.GetDirectoryName(Path.GetFullPath(path));
            if (string.IsNullOrEmpty(directory)) {
                return;
            }

            if (!Directory.Exists(directory)) {
                Directory.CreateDirectory(directory);
                return;
            }

            var directoryInfo = new DirectoryInfo(directory);
            if (directoryInfo.Attributes.HasFlag(FileAttributes.ReadOnly)) {
                throw new IOException($"Failed to save to '{path}'. The directory is read-only.");
            }
        }

        /// <summary>
        /// Saves the document and optionally opens it.
        /// </summary>
        /// <param name="filePath">Path to save to.</param>
        /// <param name="openExcel">Whether to open the file after saving.</param>
        public void Save(string filePath, bool openExcel) {
            Save(filePath, openExcel, options: null);
        }

        /// <summary>
        /// Saves the document with optional robustness options.
        /// </summary>
        /// <param name="filePath">Destination path. When empty, uses the original <see cref="FilePath"/>.</param>
        /// <param name="openExcel">When true, opens the saved file in the system's associated app.</param>
        /// <param name="options">Optional save behaviors (safe defined-name repair, post-save Open XML validation).</param>
        public void Save(string filePath, bool openExcel, ExcelSaveOptions? options) {
            if (string.IsNullOrEmpty(filePath) && string.IsNullOrEmpty(FilePath)) {
                if (_sourceStream != null) {
                    Save(_sourceStream, options);
                    return;
                }

                throw new InvalidOperationException("This workbook is not associated with a file path. Provide a file path or call Save(Stream).");
            }

            var path = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
            var originalFilePath = FilePath;

            // Ensure target directory is writable
            if (File.Exists(path) && new FileInfo(path).IsReadOnly) {
                throw new IOException($"Failed to save to '{path}'. The file is read-only.");
            }
            EnsureDirectoryWritable(path);

            if (TrySaveDirectDataSetPackageToFile(path, options, CancellationToken.None, out _)) {
                if (openExcel) {
                    Helpers.Open(path, true);
                }

                return;
            }

            bool preferExtendedPackageWriter = _materializedDirectDataSetFastSaveModel != null;
            Stopwatch? saveStageWatch = Execution.OnTiming == null ? null : Stopwatch.StartNew();
            PrepareWorkbookForSave(options, skipDirectFastSaveSheetPreparation: preferExtendedPackageWriter);
            ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbook");

            string? extendedPackageSkipReason = null;
            if (preferExtendedPackageWriter
                && TrySaveWithExtendedPackageToFile(path, options, out extendedPackageSkipReason)) {
                if (openExcel) {
                    Helpers.Open(path, true);
                }

                return;
            }

            if (preferExtendedPackageWriter) {
                PrepareWorkbookForSave(options);
                ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbookFallback");
            }

            if (TrySaveWithSimplePackageToFile(path, options, out string? fastPackageSkipReason, alreadyPrepared: true)) {
                if (openExcel) {
                    Helpers.Open(path, true);
                }

                return;
            }

            if (!preferExtendedPackageWriter
                && TrySaveWithExtendedPackageToFile(path, options, out extendedPackageSkipReason)) {
                if (openExcel) {
                    Helpers.Open(path, true);
                }

                return;
            }

            var payload = PreparePackageForSave(options);
            try {
                var finalizedBytes = FinalizePackageBytes(payload);
                ThrowIfOpenXmlValidationFails(finalizedBytes, options);
                CommitPreparedPackageToFile(path, finalizedBytes);
                ReloadFromBytes(finalizedBytes);
                FilePath = path;
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(extendedPackageSkipReason ?? fastPackageSkipReason);

                if (openExcel) {
                    Helpers.Open(path, true);
                }
            } catch {
                TryRestoreDocumentState(payload);
                FilePath = originalFilePath;
                throw;
            }
        }

        /// <summary>
        /// Saves the workbook as a password-encrypted Office Open XML package.
        /// </summary>
        /// <param name="filePath">Destination path. When empty, uses the original <see cref="FilePath"/>.</param>
        /// <param name="password">Password used to encrypt the workbook package.</param>
        /// <param name="openExcel">When true, opens the saved file in the system's associated app.</param>
        /// <param name="saveOptions">Optional save behaviors (safe defined-name repair, post-save Open XML validation).</param>
        public void SaveEncrypted(string filePath, string password, bool openExcel = false, ExcelSaveOptions? saveOptions = null) {
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (string.IsNullOrEmpty(filePath) && string.IsNullOrEmpty(FilePath)) {
                throw new InvalidOperationException("This workbook is not associated with a file path. Provide a file path or call SaveEncrypted(Stream, ...).");
            }

            var path = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
            var originalFilePath = FilePath;
            if (File.Exists(path) && new FileInfo(path).IsReadOnly) {
                throw new IOException($"Failed to save to '{path}'. The file is read-only.");
            }
            EnsureDirectoryWritable(path);

            var payload = PreparePackageForSave(saveOptions);
            try {
                var finalizedBytes = FinalizePackageBytes(payload);
                ThrowIfOpenXmlValidationFails(finalizedBytes, saveOptions);
                var encryptedBytes = OfficeEncryption.EncryptPackage(finalizedBytes, password);
                CommitPreparedPackageToFile(path, encryptedBytes);
                ReloadFromBytes(finalizedBytes);
                FilePath = path;
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Encrypted saves use the standard package finalization path.");

                if (openExcel) {
                    Helpers.Open(path, true);
                }
            } catch {
                TryRestoreDocumentState(payload);
                FilePath = originalFilePath;
                throw;
            }
        }

        /// <summary>
        /// Saves the document and writes an optional OpenXML validation report (sidecar file)
        /// next to the saved .xlsx when issues are detected. Useful to diagnose any remaining
        /// problems that could cause Excel's repair dialog.
        /// </summary>
        /// <param name="filePath">Destination path. Empty uses <see cref="FilePath"/>.</param>
        /// <param name="openExcel">When true, launches the saved file.</param>
        /// <param name="writeReportOnIssues">When true (default), writes <c>.xlsx.validation.txt</c> on issues.</param>
        public void SafeSave(string filePath = "", bool openExcel = false, bool writeReportOnIssues = true) {
            Save(filePath, openExcel);
            try {
                var errs = ValidateDocument();
                if (errs.Count > 0 && writeReportOnIssues) {
                    var target = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
                    var reportPath = System.IO.Path.ChangeExtension(target, ".xlsx.validation.txt");
                    var lines = new System.Collections.Generic.List<string>(errs.Count);
                    foreach (var e in errs) {
                        lines.Add($"{e.ErrorType}: {e.Description} at {e.Path?.XPath}");
                    }
                    System.IO.File.WriteAllLines(reportPath, lines);
                }
            } catch { }
        }

        /// <summary>
        /// Saves the document without opening it.
        /// </summary>
        public void Save() {
            this.Save("", false);
        }

        /// <summary>
        /// Saves the document and optionally opens it.
        /// </summary>
        /// <param name="openExcel">Whether to open the file after saving.</param>
        public void Save(bool openExcel) {
            this.Save("", openExcel);
        }

        /// <summary>
        /// Fluent sugar: compose a worksheet using <see cref="Fluent.SheetComposer"/> without exposing the builder type to callers.
        /// </summary>
        public void Compose(string sheetName, System.Action<OfficeIMO.Excel.Fluent.SheetComposer> compose, OfficeIMO.Excel.Fluent.SheetTheme? theme = null) {
            if (compose == null) throw new System.ArgumentNullException(nameof(compose));
            var c = new OfficeIMO.Excel.Fluent.SheetComposer(this, sheetName, theme);
            compose(c);
        }

        /// <summary>
        /// Asynchronously saves the document.
        /// </summary>
        /// <param name="filePath">Optional path to save to.</param>
        /// <param name="openExcel">Whether to open Excel after saving.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public async Task SaveAsync(string filePath, bool openExcel, CancellationToken cancellationToken = default) {
            await SaveAsync(filePath, openExcel, options: null, cancellationToken: cancellationToken).ConfigureAwait(false);
        }

        /// <summary>
        /// Asynchronously saves the document with optional robustness options.
        /// </summary>
        /// <param name="filePath">Destination path. When empty, uses the original <see cref="FilePath"/>.</param>
        /// <param name="openExcel">When true, opens the saved file in the system's associated app.</param>
        /// <param name="options">Optional save behaviors (safe defined-name repair, post-save Open XML validation).</param>
        /// <param name="cancellationToken">Cancels the asynchronous save work.</param>
        public async Task SaveAsync(string filePath, bool openExcel, ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            if (string.IsNullOrEmpty(filePath) && string.IsNullOrEmpty(FilePath)) {
                if (_sourceStream != null) {
                    await SaveAsync(_sourceStream, options, cancellationToken).ConfigureAwait(false);
                    return;
                }

                throw new InvalidOperationException("This workbook is not associated with a file path. Provide a file path or call Save(Stream).");
            }

            var target = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
            var originalFilePath = FilePath;
            if (File.Exists(target) && new FileInfo(target).IsReadOnly) {
                throw new IOException($"Failed to save to '{target}'. The file is read-only.");
            }
            EnsureDirectoryWritable(target);

            if (TrySaveDirectDataSetPackageToFile(target, options, cancellationToken, out _)) {
                if (openExcel) {
                    Open(target, true);
                }

                return;
            }

            cancellationToken.ThrowIfCancellationRequested();
            bool preferExtendedPackageWriter = _materializedDirectDataSetFastSaveModel != null;
            Stopwatch? saveStageWatch = Execution.OnTiming == null ? null : Stopwatch.StartNew();
            PrepareWorkbookForSave(options, skipDirectFastSaveSheetPreparation: preferExtendedPackageWriter);
            ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbook");

            string? extendedPackageSkipReason = null;
            if (preferExtendedPackageWriter
                && TrySaveWithExtendedPackageToFile(target, options, out extendedPackageSkipReason, cancellationToken)) {
                if (openExcel) {
                    Open(target, true);
                }

                return;
            }

            if (preferExtendedPackageWriter) {
                PrepareWorkbookForSave(options);
                ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbookFallback");
            }

            if (TrySaveWithSimplePackageToFile(target, options, out string? fastPackageSkipReason, cancellationToken, alreadyPrepared: true)) {
                if (openExcel) {
                    Open(target, true);
                }

                return;
            }

            if (!preferExtendedPackageWriter
                && TrySaveWithExtendedPackageToFile(target, options, out extendedPackageSkipReason, cancellationToken)) {
                if (openExcel) {
                    Open(target, true);
                }

                return;
            }

            var payload = PreparePackageForSave(options);
            try {
                var finalizedBytes = FinalizePackageBytes(payload);
                ThrowIfOpenXmlValidationFails(finalizedBytes, options);
                await CommitPreparedPackageToFileAsync(target, finalizedBytes, cancellationToken).ConfigureAwait(false);
                ReloadFromBytes(finalizedBytes);
                FilePath = target;
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(extendedPackageSkipReason ?? fastPackageSkipReason);

                if (openExcel) {
                    Open(target, true);
                }
            } catch {
                TryRestoreDocumentState(payload);
                FilePath = originalFilePath;
                throw;
            }
        }

        /// <summary>
        /// Saves the document into a writable stream.
        /// </summary>
        /// <param name="destination">Writable stream that receives the Excel package content.</param>
        public void Save(Stream destination) {
            Save(destination, options: null);
        }

        /// <summary>
        /// Saves the document into a writable stream with optional robustness options.
        /// </summary>
        /// <param name="destination">Writable stream that receives the Excel package content.</param>
        /// <param name="options">Optional save behaviors (safe defined-name repair, post-save Open XML validation).</param>
        public void Save(Stream destination, ExcelSaveOptions? options) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));

            if (TryWriteUnchangedPackageToStream(destination, options)) {
                LastSaveDiagnostics = ExcelSaveDiagnostics.UnchangedPackage();
                return;
            }

            if (TryWriteDirectDataSetPackage(destination, options, updateDocumentState: true, CancellationToken.None, out _)) {
                LastSaveDiagnostics = ExcelSaveDiagnostics.DirectDataSetPackage();
                return;
            }

            bool preferExtendedPackageWriter = _materializedDirectDataSetFastSaveModel != null;
            Stopwatch? saveStageWatch = Execution.OnTiming == null ? null : Stopwatch.StartNew();
            PrepareWorkbookForSave(options, skipDirectFastSaveSheetPreparation: preferExtendedPackageWriter);
            ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbook");

            string? extendedPackageSkipReason = null;
            if (preferExtendedPackageWriter
                && TryWriteExtendedWorkbookPackage(destination, options, updateDocumentState: true, out extendedPackageSkipReason)) {
                LastSaveDiagnostics = ExcelSaveDiagnostics.ExtendedPackage();
                return;
            }

            if (preferExtendedPackageWriter) {
                PrepareWorkbookForSave(options);
                ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbookFallback");
            }

            if (TryWriteSimpleWorkbookPackage(destination, options, updateDocumentState: true, out string? fastPackageSkipReason)) {
                LastSaveDiagnostics = ExcelSaveDiagnostics.SimplePackage();
                return;
            }

            if (!preferExtendedPackageWriter
                && TryWriteExtendedWorkbookPackage(destination, options, updateDocumentState: true, out extendedPackageSkipReason)) {
                LastSaveDiagnostics = ExcelSaveDiagnostics.ExtendedPackage();
                return;
            }

            var payload = PreparePackageForSave(options, closeDocument: false);
            try {
                var finalizedBytes = FinalizePackageBytes(payload);
                ThrowIfOpenXmlValidationFails(finalizedBytes, options);
                PrepareDestinationStreamForWrite(destination);
                destination.Write(finalizedBytes, 0, finalizedBytes.Length);
                try { destination.Flush(); } catch (NotSupportedException) { }
                MarkPackageClean(finalizedBytes);
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(extendedPackageSkipReason ?? fastPackageSkipReason);
            } catch {
                TryRestoreDocumentState(payload);
                throw;
            }
        }

        /// <summary>
        /// Saves the workbook as a password-encrypted Office Open XML package to a stream.
        /// </summary>
        /// <param name="destination">Writable stream that receives the encrypted workbook.</param>
        /// <param name="password">Password used to encrypt the workbook package.</param>
        /// <param name="saveOptions">Optional save behaviors (safe defined-name repair, post-save Open XML validation).</param>
        public void SaveEncrypted(Stream destination, string password, ExcelSaveOptions? saveOptions = null) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));

            if (CanUseUnchangedPackageFastPath(saveOptions) && _unchangedPackageBytes != null) {
                OfficeEncryption.EncryptPackageToStream(_unchangedPackageBytes, password, destination);
                LastSaveDiagnostics = ExcelSaveDiagnostics.UnchangedPackage();
                return;
            }

            var payload = PreparePackageForSave(saveOptions, closeDocument: false);
            try {
                var finalizedBytes = FinalizePackageBytes(payload);
                ThrowIfOpenXmlValidationFails(finalizedBytes, saveOptions);
                OfficeEncryption.EncryptPackageToStream(finalizedBytes, password, destination);
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard("Encrypted saves use the standard package finalization path.");
            } catch {
                TryRestoreDocumentState(payload);
                throw;
            }
        }

        /// <summary>
        /// Asynchronously saves the document into a writable stream.
        /// </summary>
        /// <param name="destination">Writable stream that receives the Excel package content.</param>
        /// <param name="cancellationToken">Cancels the asynchronous save work.</param>
        public Task SaveAsync(Stream destination, CancellationToken cancellationToken = default) {
            return SaveAsync(destination, options: null, cancellationToken);
        }

        /// <summary>
        /// Asynchronously saves the document into a writable stream with optional robustness options.
        /// </summary>
        /// <param name="destination">Writable stream that receives the Excel package content.</param>
        /// <param name="options">Optional save behaviors (safe defined-name repair, post-save Open XML validation).</param>
        /// <param name="cancellationToken">Cancels the asynchronous save work.</param>
        public async Task SaveAsync(Stream destination, ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));

            if (await TryWriteUnchangedPackageToStreamAsync(destination, options, cancellationToken).ConfigureAwait(false)) {
                LastSaveDiagnostics = ExcelSaveDiagnostics.UnchangedPackage();
                return;
            }

            if (TryWriteDirectDataSetPackage(destination, options, updateDocumentState: true, cancellationToken, out _)) {
                LastSaveDiagnostics = ExcelSaveDiagnostics.DirectDataSetPackage();
                return;
            }

            bool preferExtendedPackageWriter = _materializedDirectDataSetFastSaveModel != null;
            Stopwatch? saveStageWatch = Execution.OnTiming == null ? null : Stopwatch.StartNew();
            PrepareWorkbookForSave(options, skipDirectFastSaveSheetPreparation: preferExtendedPackageWriter);
            ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbook");
            cancellationToken.ThrowIfCancellationRequested();

            string? extendedPackageSkipReason = null;
            if (preferExtendedPackageWriter
                && TryWriteExtendedWorkbookPackage(destination, options, updateDocumentState: true, out extendedPackageSkipReason, cancellationToken)) {
                LastSaveDiagnostics = ExcelSaveDiagnostics.ExtendedPackage();
                return;
            }

            if (preferExtendedPackageWriter) {
                PrepareWorkbookForSave(options);
                ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbookFallback");
            }

            if (TryWriteSimpleWorkbookPackage(destination, options, updateDocumentState: true, out string? fastPackageSkipReason, cancellationToken)) {
                LastSaveDiagnostics = ExcelSaveDiagnostics.SimplePackage();
                return;
            }

            if (!preferExtendedPackageWriter
                && TryWriteExtendedWorkbookPackage(destination, options, updateDocumentState: true, out extendedPackageSkipReason, cancellationToken)) {
                LastSaveDiagnostics = ExcelSaveDiagnostics.ExtendedPackage();
                return;
            }

            var payload = PreparePackageForSave(options, closeDocument: false);
            try {
                var finalizedBytes = FinalizePackageBytes(payload);
                ThrowIfOpenXmlValidationFails(finalizedBytes, options);
                PrepareDestinationStreamForWrite(destination);
                await destination.WriteAsync(finalizedBytes, 0, finalizedBytes.Length, cancellationToken).ConfigureAwait(false);
                try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }
                MarkPackageClean(finalizedBytes);
                LastSaveDiagnostics = ExcelSaveDiagnostics.Standard(extendedPackageSkipReason ?? fastPackageSkipReason);
            } catch {
                TryRestoreDocumentState(payload);
                throw;
            }
        }

        /// <summary>
        /// Asynchronously saves the document.
        /// </summary>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(CancellationToken cancellationToken = default) {
            return SaveAsync("", false, cancellationToken);
        }


        /// <summary>
        /// Asynchronously saves the document and optionally opens Excel.
        /// </summary>
        /// <param name="openExcel">Whether to open Excel after saving.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public Task SaveAsync(bool openExcel, CancellationToken cancellationToken = default) {
            return SaveAsync("", openExcel, cancellationToken);
        }
    }
}
