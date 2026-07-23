using OfficeIMO.Drawing.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.Utilities;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        /// <summary>Opens the associated workbook in the operating system's registered application.</summary>
        public void OpenInApplication(string? filePath = null) {
            string? target = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
            if (string.IsNullOrEmpty(target)) {
                throw new InvalidOperationException("The workbook has no associated file path.");
            }
            OfficeFileLauncher.Open(target!);
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

        private static void EnsureDestinationFileWritable(string path) {
            if (File.Exists(path) && new FileInfo(path).IsReadOnly) {
                throw new IOException($"Failed to save to '{path}'. The file is read-only.");
            }
        }

        private void EnsureWritableForSave() {
            if (_spreadSheetDocument.FileOpenAccess == FileAccess.Read) {
                throw new InvalidOperationException("The workbook is read-only and cannot be saved.");
            }
        }

        /// <summary>
        /// Saves the document without opening it.
        /// </summary>
        /// <param name="filePath">Path to save to.</param>
        public void Save(string filePath) {
            SaveFileCore(filePath, options: null);
        }

        /// <summary>Saves the document with typed options and no positional Boolean.</summary>
        /// <param name="filePath">Path to save to.</param>
        /// <param name="options">Optional save policy settings.</param>
        public void Save(string filePath, ExcelSaveOptions? options) {
            SaveFileCore(filePath, options);
        }

        private void SaveFileCore(string? filePath, ExcelSaveOptions? options) {
            EnsureWritableForSave();
            if (string.IsNullOrEmpty(filePath) && string.IsNullOrEmpty(FilePath)) {
                if (_sourceStream != null) {
                    Save(_sourceStream, options);
                    return;
                }

                throw new InvalidOperationException("This workbook is not associated with a file path. Provide a file path or call Save(Stream).");
            }

            string path = (string.IsNullOrEmpty(filePath) ? FilePath : filePath)
                ?? throw new InvalidOperationException("This workbook is not associated with a file path. Provide a file path or call Save(Stream).");
            var originalFilePath = FilePath;
            EnsureLegacyXlsSaveDoesNotDropImportedContent(
                options,
                preserveLinkedVbaProject: ExcelDocumentLoadRouting.HasLegacyXlsExtension(path));
            EnsureLegacyBinaryExcelSaveTargetSupported(path, allowNativeXls: true, options);
            EnsureXlsbFileTargetSupported(path, options);

            // Ensure target directory is writable
            EnsureDestinationFileWritable(path);
            EnsureDirectoryWritable(path);

            if (TrySaveUnchangedXlsbToFile(path, options)) {
                return;
            }

            if (TrySaveNativeLegacyXlsToFile(path, options)) {
                return;
            }

            AlignSpreadsheetDocumentTypeWithFilePath(path);

            if (TrySaveDirectDataSetPackageToFile(path, options, CancellationToken.None, out _)) {
                return;
            }

            bool preferExtendedPackageWriter = _materializedDirectDataSetFastSaveModel != null;
            Stopwatch? saveStageWatch = Execution.OnTiming == null ? null : Stopwatch.StartNew();
            PrepareWorkbookForSave(options, skipDirectFastSaveSheetPreparation: preferExtendedPackageWriter);
            ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbook");

            string? extendedPackageSkipReason = null;
            if (preferExtendedPackageWriter
                && TrySaveWithExtendedPackageToFile(path, options, out extendedPackageSkipReason)) {
                return;
            }

            if (preferExtendedPackageWriter) {
                MaterializeDirectDataSetFastSaveModelIfNeeded();
                PrepareWorkbookForSave(options);
                ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbookFallback");
            }

            if (TrySaveWithSimplePackageToFile(path, options, out string? fastPackageSkipReason, alreadyPrepared: true)) {
                return;
            }

            if (!preferExtendedPackageWriter
                && TrySaveWithExtendedPackageToFile(path, options, out extendedPackageSkipReason)) {
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
        /// <param name="saveOptions">Optional save behaviors (safe defined-name repair, post-save Open XML validation).</param>
        public void SaveEncrypted(string filePath, string password, ExcelSaveOptions? saveOptions = null) {
            if (password == null) throw new ArgumentNullException(nameof(password));
            EnsureWritableForSave();
            if (string.IsNullOrEmpty(filePath) && string.IsNullOrEmpty(FilePath)) {
                throw new InvalidOperationException("This workbook is not associated with a file path. Provide a file path or call SaveEncrypted(Stream, ...).");
            }

            string path = string.IsNullOrEmpty(filePath) ? FilePath! : filePath;
            var originalFilePath = FilePath;
            EnsureLegacyXlsSaveDoesNotDropImportedContent(saveOptions);
            EnsureLegacyBinaryEncryptedSaveTargetSupported(path);
            EnsureXlsbFileTargetSupported(path, saveOptions, allowUnchangedCopy: false);
            AlignSpreadsheetDocumentTypeWithFilePath(path);
            EnsureDestinationFileWritable(path);
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
            } catch {
                TryRestoreDocumentState(payload);
                FilePath = originalFilePath;
                throw;
            }
        }

        /// <summary>
        /// Saves the document without opening it.
        /// </summary>
        public void Save() {
            Save(options: null);
        }

        /// <summary>Saves to the associated destination with optional save settings.</summary>
        public void Save(ExcelSaveOptions? options) {
            if (string.IsNullOrEmpty(FilePath) && _sourceStream != null) {
                Save(_sourceStream, options);
            } else {
                SaveFileCore(FilePath, options);
            }
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
        /// <param name="options">Optional save policy settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        private async Task SaveFileAsyncCore(string? filePath, ExcelSaveOptions? options, CancellationToken cancellationToken) {
            EnsureWritableForSave();
            if (string.IsNullOrEmpty(filePath) && string.IsNullOrEmpty(FilePath)) {
                if (_sourceStream != null) {
                    await SaveAsync(_sourceStream, options, cancellationToken).ConfigureAwait(false);
                    return;
                }

                throw new InvalidOperationException("This workbook is not associated with a file path. Provide a file path or call Save(Stream).");
            }

            string target = (string.IsNullOrEmpty(filePath) ? FilePath : filePath)
                ?? throw new InvalidOperationException("This workbook is not associated with a file path. Provide a file path or call Save(Stream).");
            var originalFilePath = FilePath;
            EnsureLegacyXlsSaveDoesNotDropImportedContent(
                options,
                preserveLinkedVbaProject: ExcelDocumentLoadRouting.HasLegacyXlsExtension(target));
            EnsureLegacyBinaryExcelSaveTargetSupported(target, allowNativeXls: true, options);
            EnsureXlsbFileTargetSupported(target, options);
            EnsureDestinationFileWritable(target);
            EnsureDirectoryWritable(target);

            if (await TrySaveUnchangedXlsbToFileAsync(target, options, cancellationToken).ConfigureAwait(false)) {
                return;
            }

            if (await TrySaveNativeLegacyXlsToFileAsync(target, options, cancellationToken).ConfigureAwait(false)) {
                return;
            }

            AlignSpreadsheetDocumentTypeWithFilePath(target);

            if (TrySaveDirectDataSetPackageToFile(target, options, cancellationToken, out _)) {
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
                return;
            }

            if (preferExtendedPackageWriter) {
                MaterializeDirectDataSetFastSaveModelIfNeeded();
                PrepareWorkbookForSave(options);
                ReportSaveTiming(saveStageWatch, "Save.PrepareWorkbookFallback");
            }

            if (TrySaveWithSimplePackageToFile(target, options, out string? fastPackageSkipReason, cancellationToken, alreadyPrepared: true)) {
                return;
            }

            if (!preferExtendedPackageWriter
                && TrySaveWithExtendedPackageToFile(target, options, out extendedPackageSkipReason, cancellationToken)) {
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
            } catch {
                TryRestoreDocumentState(payload);
                FilePath = originalFilePath;
                throw;
            }
        }

        /// <summary>Encodes the workbook in the selected physical format.</summary>
        public byte[] ToBytes(ExcelFileFormat format = ExcelFileFormat.Xlsx, ExcelSaveOptions? options = null) {
            using var stream = new MemoryStream();
            Save(stream, format, options);
            return stream.ToArray();
        }

        /// <summary>Encodes the workbook in a new writable memory stream positioned at the beginning.</summary>
        public MemoryStream ToStream(ExcelFileFormat format = ExcelFileFormat.Xlsx, ExcelSaveOptions? options = null) =>
            new MemoryStream(ToBytes(format, options));

        /// <summary>
        /// Saves the document into a writable stream as XLSX.
        /// </summary>
        /// <param name="destination">Writable stream that receives the Excel package content.</param>
        public void Save(Stream destination) {
            Save(destination, ExcelFileFormat.Xlsx, options: null);
        }

        /// <summary>
        /// Saves the document into a writable stream with optional robustness options.
        /// </summary>
        /// <param name="destination">Writable stream that receives the Excel package content.</param>
        /// <param name="options">Optional save behaviors (safe defined-name repair, post-save Open XML validation).</param>
        public void Save(Stream destination, ExcelSaveOptions? options) {
            Save(destination, ExcelFileFormat.Xlsx, options);
        }

        /// <summary>Saves the workbook to a stream in the explicitly selected physical format.</summary>
        /// <param name="destination">Writable destination stream. This one-time save does not change the associated destination.</param>
        /// <param name="format">Physical XLSX, XLS, or XLSB format.</param>
        /// <param name="options">Optional save settings.</param>
        public void Save(Stream destination, ExcelFileFormat format, ExcelSaveOptions? options = null) {
            SaveToStreamCore(destination, format, options);
            if (destination.CanSeek) destination.Seek(0, SeekOrigin.Begin);
        }

        private void SaveToStreamCore(Stream destination, ExcelFileFormat format, ExcelSaveOptions? options) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
            EnsureXlsbStreamTargetSupported(format, options);
            EnsureWritableForSave();
            EnsureLegacyXlsSaveDoesNotDropImportedContent(
                options,
                preserveLinkedVbaProject: format == ExcelFileFormat.Xls);

            if (TrySaveUnchangedXlsbToStream(destination, format, options)) {
                return;
            }

            if (TrySaveNativeLegacyXlsToStream(destination, format, options)) {
                return;
            }

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
                MaterializeDirectDataSetFastSaveModelIfNeeded();
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
            EnsureWritableForSave();
            EnsureLegacyXlsSaveDoesNotDropImportedContent(saveOptions);

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
            return SaveAsync(destination, ExcelFileFormat.Xlsx, options: null, cancellationToken);
        }

        /// <summary>
        /// Asynchronously saves the document into a writable stream with optional robustness options.
        /// </summary>
        /// <param name="destination">Writable stream that receives the Excel package content.</param>
        /// <param name="options">Optional save behaviors (safe defined-name repair, post-save Open XML validation).</param>
        /// <param name="cancellationToken">Cancels the asynchronous save work.</param>
        public Task SaveAsync(Stream destination, ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            return SaveAsync(destination, ExcelFileFormat.Xlsx, options, cancellationToken);
        }

        /// <summary>Asynchronously saves the workbook to a stream in the selected physical format.</summary>
        /// <param name="destination">Writable destination stream.</param>
        /// <param name="format">Physical XLSX or XLS format.</param>
        /// <param name="options">Optional save settings.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        public async Task SaveAsync(Stream destination, ExcelFileFormat format, ExcelSaveOptions? options = null, CancellationToken cancellationToken = default) {
            await SaveToStreamAsyncCore(destination, format, options, cancellationToken).ConfigureAwait(false);
            if (destination.CanSeek) destination.Seek(0, SeekOrigin.Begin);
        }

        private async Task SaveToStreamAsyncCore(Stream destination, ExcelFileFormat format, ExcelSaveOptions? options, CancellationToken cancellationToken) {
            if (destination == null) throw new ArgumentNullException(nameof(destination));
            if (!destination.CanWrite) throw new ArgumentException("Destination stream must be writable.", nameof(destination));
            EnsureXlsbStreamTargetSupported(format, options);
            EnsureWritableForSave();
            EnsureLegacyXlsSaveDoesNotDropImportedContent(
                options,
                preserveLinkedVbaProject: format == ExcelFileFormat.Xls);

            if (await TrySaveUnchangedXlsbToStreamAsync(destination, format, options, cancellationToken).ConfigureAwait(false)) {
                return;
            }

            if (await TrySaveNativeLegacyXlsToStreamAsync(destination, format, options, cancellationToken).ConfigureAwait(false)) {
                return;
            }

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
                MaterializeDirectDataSetFastSaveModelIfNeeded();
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
            return SaveAsync(options: null, cancellationToken);
        }

        /// <summary>Asynchronously saves to the associated destination with optional save settings.</summary>
        public Task SaveAsync(ExcelSaveOptions? options, CancellationToken cancellationToken = default) {
            if (string.IsNullOrEmpty(FilePath) && _sourceStream != null) {
                return SaveAsync(_sourceStream, options, cancellationToken);
            }
            return SaveFileAsyncCore(FilePath, options, cancellationToken);
        }

        /// <summary>Asynchronously saves to a path with optional save settings.</summary>
        public Task SaveAsync(string filePath, ExcelSaveOptions? options = null, CancellationToken cancellationToken = default) {
            return SaveFileAsyncCore(filePath, options, cancellationToken);
        }
    }
}
