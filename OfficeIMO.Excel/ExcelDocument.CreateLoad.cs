using OfficeIMO.Drawing.Internal;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
using OfficeIMO.Excel.Xlsb;
using OfficeIMO.Excel.Xlsb.Model;
using OfficeIMO.Excel.Xlsb.Projection;
using OfficeIMO.Excel.Utilities;
using OfficeIMO.Drawing;
using System.IO.Packaging;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using System;
using System.Diagnostics;
using System.IO;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

        /// <summary>Creates a detached workbook that must be saved explicitly to a path or stream.</summary>
        /// <param name="options">Creation options. SaveOnDispose is invalid without an associated destination.</param>
        public static ExcelDocument Create(ExcelCreateOptions? options = null) {
            ExcelCreateOptions resolved = options ?? new ExcelCreateOptions();
            if (resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                throw new ArgumentException("SaveOnDispose requires an associated file path or writable stream.", nameof(options));
            }

            var packageStream = new MemoryStream(StreamBufferSize);
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(packageStream, resolved.DocumentType, autoSave: false);
            return CreateNewDocument(
                spreadSheetDocument,
                filePath: null,
                packageStream,
                sourceStream: null,
                resolved.PersistenceMode,
                copyPackageToSourceOnDispose: false,
                leaveSourceStreamOpen: true);
        }

        /// <summary>
        /// Creates a new Excel document at the specified path.
        /// </summary>
        /// <param name="filePath">Path to the new file.</param>
        /// <param name="options">Creation and persistence options.</param>
        /// <returns>Created <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Create(string filePath, ExcelCreateOptions? options = null) {
            if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("File path cannot be empty.", nameof(filePath));
            ExcelCreateOptions resolved = options ?? new ExcelCreateOptions();
            bool saveOnDispose = resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose;
            if (saveOnDispose && string.IsNullOrEmpty(filePath)) {
                throw new ArgumentException("SaveOnDispose requires an associated file path or writable stream.", nameof(filePath));
            }

            Stream packageStream = saveOnDispose
                ? new NonDisposingMemoryStream(StreamBufferSize)
                : new MemoryStream(StreamBufferSize);
            SpreadsheetDocumentType documentType = ResolveSpreadsheetDocumentType(filePath) ?? resolved.DocumentType;
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(packageStream, documentType, autoSave: false);
            return CreateNewDocument(
                spreadSheetDocument,
                filePath,
                packageStream,
                sourceStream: null,
                resolved.PersistenceMode,
                copyPackageToSourceOnDispose: false,
                leaveSourceStreamOpen: true,
                copyPackageToFilePathOnDispose: saveOnDispose);
        }

        /// <summary>
        /// Creates a new Excel document in memory and optionally persists it to the provided stream on dispose.
        /// </summary>
        /// <param name="stream">Destination stream to receive the workbook package.</param>
        /// <param name="options">Creation and persistence options.</param>
        /// <returns>Created <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Create(Stream stream, ExcelCreateOptions? options = null) {
            OfficeDocumentLifecycle.EnsureAssociatedDestination(stream, nameof(stream));
            ExcelCreateOptions resolved = options ?? new ExcelCreateOptions();
            bool saveOnDispose = resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose;

            Stream packageStream = saveOnDispose
                ? new NonDisposingMemoryStream(StreamBufferSize)
                : new MemoryStream(StreamBufferSize);

            var spreadSheetDocument = SpreadsheetDocument.Create(packageStream, resolved.DocumentType, false);
            return CreateNewDocument(spreadSheetDocument, filePath: null, packageStream, stream, resolved.PersistenceMode, saveOnDispose, leaveSourceStreamOpen: true);
        }

        private static SpreadsheetDocumentType? ResolveSpreadsheetDocumentType(string? filePath) {
            if (string.IsNullOrWhiteSpace(filePath)) {
                return null;
            }

            return Path.GetExtension(filePath).ToLowerInvariant() switch {
                ".xlsx" => SpreadsheetDocumentType.Workbook,
                ".xlsm" => SpreadsheetDocumentType.MacroEnabledWorkbook,
                ".xltx" => SpreadsheetDocumentType.Template,
                ".xltm" => SpreadsheetDocumentType.MacroEnabledTemplate,
                ".xlam" => SpreadsheetDocumentType.AddIn,
                _ => null
            };
        }

        private void AlignSpreadsheetDocumentTypeWithFilePath(string filePath) {
            SpreadsheetDocumentType? target = ResolveSpreadsheetDocumentType(filePath);
            if (!target.HasValue) {
                return;
            }

            bool macroFreeTarget = target.Value == SpreadsheetDocumentType.Workbook
                || target.Value == SpreadsheetDocumentType.Template;
            bool currentlyMacroEnabled = _spreadSheetDocument.DocumentType == SpreadsheetDocumentType.MacroEnabledWorkbook
                || _spreadSheetDocument.DocumentType == SpreadsheetDocumentType.MacroEnabledTemplate
                || _spreadSheetDocument.DocumentType == SpreadsheetDocumentType.AddIn;
            if (macroFreeTarget && HasMacros && currentlyMacroEnabled) {
                throw new InvalidOperationException(
                    "A workbook containing VBA cannot be saved to a macro-free .xlsx or .xltx destination. Use .xlsm/.xltm or remove the VBA project first.");
            }

            if (_spreadSheetDocument.DocumentType != target.Value) {
                _spreadSheetDocument.ChangeDocumentType(target.Value);
                MarkPackageDirty();
            }
        }

        private static ExcelDocument CreateNewDocument(
            SpreadsheetDocument spreadSheetDocument,
            string? filePath,
            Stream? packageStream,
            Stream? sourceStream,
            DocumentPersistenceMode persistenceMode,
            bool copyPackageToSourceOnDispose,
            bool leaveSourceStreamOpen,
            bool copyPackageToFilePathOnDispose = false,
            Stream? ownedOpenStream = null) {
            bool keepPackageStream = copyPackageToSourceOnDispose || copyPackageToFilePathOnDispose;
            var document = new ExcelDocument {
                FilePath = filePath,
                _spreadSheetDocument = spreadSheetDocument,
                _persistenceMode = persistenceMode
            };

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadSheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            document._workBookPart = workbookpart;

            document._packageStream = keepPackageStream ? packageStream : null;
            document._sourceStream = sourceStream;
            document._ownedOpenStream = ownedOpenStream;
            document._copyPackageToSourceOnDispose = copyPackageToSourceOnDispose && sourceStream != null;
            document._copyPackageToFilePathOnDispose = copyPackageToFilePathOnDispose && packageStream != null && !string.IsNullOrEmpty(filePath);
            document._leaveSourceStreamOpen = leaveSourceStreamOpen;
            document._packageContentTypesKnownNormalized = false;
            document._simplePackageContentKnown = false;
            document._requiresSavePreflight = true;
            document._packageDirty = true;
            document._packagePropertiesDirty = false;
            document._unchangedPackageBytes = null;

            // Initialize document property helpers
            document.BuiltinDocumentProperties = new BuiltinDocumentProperties(document);
            document.ApplicationProperties = new ApplicationProperties(document);
            document.CustomDocumentProperties.SetChangeHandler(document.MarkCustomDocumentPropertiesChanged);
            document.LoadCustomDocumentProperties();

            return document;
        }
        private static ExcelDocument CreateDocument(
            SpreadsheetDocument spreadSheetDocument,
            string? filePath,
            Stream? packageStream = null,
            Stream? sourceStream = null,
            bool copyPackageToSourceOnDispose = false,
            bool leaveSourceStreamOpen = true,
            bool copyPackageToFilePathOnDispose = false,
            Stream? ownedOpenStream = null,
            bool packageContentTypesKnownNormalized = false,
            byte[]? unchangedPackageBytes = null,
            DocumentPersistenceMode persistenceMode = DocumentPersistenceMode.Explicit) {
            bool keepPackageStream = copyPackageToSourceOnDispose || copyPackageToFilePathOnDispose;
            var document = new ExcelDocument {
                FilePath = filePath,
                _spreadSheetDocument = spreadSheetDocument,
                _workBookPart = GetWorkbookPartOrThrow(spreadSheetDocument),
                _packageStream = keepPackageStream ? packageStream : null,
                _sourceStream = sourceStream,
                _ownedOpenStream = ownedOpenStream,
                _copyPackageToSourceOnDispose = copyPackageToSourceOnDispose && sourceStream != null,
                _copyPackageToFilePathOnDispose = copyPackageToFilePathOnDispose && packageStream != null && !string.IsNullOrEmpty(filePath),
                _leaveSourceStreamOpen = leaveSourceStreamOpen,
                _packageContentTypesKnownNormalized = packageContentTypesKnownNormalized,
                _simplePackageContentKnown = false,
                _requiresSavePreflight = false,
                _packageDirty = false,
                _packagePropertiesDirty = false,
                _unchangedPackageBytes = packageContentTypesKnownNormalized ? unchangedPackageBytes : null,
                _persistenceMode = persistenceMode,
            };

            document.BuiltinDocumentProperties = new BuiltinDocumentProperties(document);
            document.ApplicationProperties = new ApplicationProperties(document);
            document.CustomDocumentProperties.SetChangeHandler(document.MarkCustomDocumentPropertiesChanged);
            document.LoadCustomDocumentProperties();
            ExcelChartAxisIdGenerator.Initialize(document._spreadSheetDocument);
            return document;
        }

        private static WorkbookPart GetWorkbookPartOrThrow(SpreadsheetDocument document) {
            if (document == null) throw new ArgumentNullException(nameof(document));

            var workbookPart = document.WorkbookPart;
            if (workbookPart != null) {
                return workbookPart;
            }

            workbookPart = document.GetPartsOfType<WorkbookPart>().FirstOrDefault();
            if (workbookPart != null) {
                return workbookPart;
            }

            throw new InvalidOperationException("WorkbookPart is null");
        }

        private static ExcelDocument LoadFromByteArray(
            byte[] bytes,
            ExcelLoadOptions options,
            string? filePath,
            Stream? originalStream = null,
            bool leaveOriginalStreamOpen = true) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (options == null) throw new ArgumentNullException(nameof(options));
            OfficeDocumentLifecycle.Validate(options.AccessMode, options.PersistenceMode, "workbook");
            if (options.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(bytes, options.PackageSecurity);
            }

            bool readOnly = options.AccessMode == DocumentAccessMode.ReadOnly;
            bool saveOnDispose = options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose;
            Stream? associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(
                originalStream,
                options.AccessMode);

            ExcelFileFormat detectedFormat = ExcelDocumentLoadRouting.DetectFormat(bytes, filePath);
            if (detectedFormat == ExcelFileFormat.Xls) {
                return LoadLegacyXlsFromNormalFlow(bytes, readOnly, saveOnDispose, filePath);
            }

            if (detectedFormat == ExcelFileFormat.Xlsb) {
                return LoadXlsbFromNormalFlow(bytes, readOnly, saveOnDispose, filePath, options.XlsbImportOptions);
            }

            var effectiveOpenSettings = CreateOpenSettings(options.OpenSettings);
            bool shouldCopyBack = saveOnDispose && associatedStream != null;
            bool shouldCopyBackToFilePath = !shouldCopyBack && !string.IsNullOrEmpty(filePath) && saveOnDispose;
            bool shouldRetainPackageStream = shouldCopyBack || shouldCopyBackToFilePath;

            MemoryStream? normalizedStream = null;

            try {
                normalizedStream = shouldRetainPackageStream
                    ? new NonDisposingMemoryStream(bytes.Length + StreamBufferSize)
                    : new MemoryStream(bytes.Length + StreamBufferSize);
                normalizedStream.Write(bytes, 0, bytes.Length);
                normalizedStream.Position = 0;

                bool normalizedContentTypes = Utilities.ExcelPackageUtilities.NormalizeContentTypes(normalizedStream, leaveOpen: true);
                normalizedStream.Position = 0;
                byte[] unchangedPackageBytes = normalizedContentTypes ? normalizedStream.ToArray() : bytes;

                var memDoc = SpreadsheetDocument.Open(normalizedStream, !readOnly, effectiveOpenSettings);
                ExcelDocument document = CreateDocument(
                    memDoc,
                    filePath,
                    shouldRetainPackageStream ? normalizedStream : null,
                    associatedStream,
                    shouldCopyBack,
                    leaveOriginalStreamOpen,
                    copyPackageToFilePathOnDispose: shouldCopyBackToFilePath,
                    packageContentTypesKnownNormalized: true,
                    unchangedPackageBytes: unchangedPackageBytes,
                    persistenceMode: options.PersistenceMode);
                document._openXmlOriginalPackageBytes = OfficeCompatibilitySourceCarrier.ContainsPackageCarrier(bytes)
                    ? (byte[])bytes.Clone()
                    : null;
                return document;
            } catch (Exception ex) when (ex is InvalidDataException || ex is OpenXmlPackageException || ex is XmlException) {
                normalizedStream?.Dispose();
                var contextMessage = filePath != null
                    ? $"Failed to open '{filePath}' after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'."
                    : "Failed to open workbook stream after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.";
                throw new IOException($"{contextMessage} See inner exception for details.", ex);
            } catch {
                DisposeStream(normalizedStream);
                throw;
            }
        }

        private static byte[] ReadAllBytes(Stream stream, OfficePackageSecurityOptions? securityOptions = null) {
            return securityOptions == null
                ? OfficeStreamReader.ReadAllBytes(stream)
                : OfficePackageSecurityInspector.ReadBounded(stream, securityOptions);
        }

        private static async Task<byte[]> ReadAllBytesAsync(Stream stream, CancellationToken cancellationToken,
            OfficePackageSecurityOptions? securityOptions = null) {
            return securityOptions == null
                ? await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false)
                : await OfficePackageSecurityInspector.ReadBoundedAsync(stream, securityOptions, cancellationToken)
                    .ConfigureAwait(false);
        }

        private static void DisposeStream(Stream? stream) {
            if (stream == null) {
                return;
            }

            if (stream is NonDisposingMemoryStream ndms) {
                ndms.DisposeUnderlying();
            } else {
                stream.Dispose();
            }
        }

        private static ExcelDocument LoadLegacyXlsFromNormalFlow(
            byte[] bytes,
            bool readOnly,
            bool saveOnDispose,
            string? filePath,
            LegacyXlsImportOptions? importOptions = null) {
            if (!readOnly && saveOnDispose) {
                throw new NotSupportedException("SaveOnDispose is not supported when loading legacy binary .xls files. Save to a new .xlsx path instead.");
            }

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(bytes, importOptions ?? new LegacyXlsImportOptions());
            LegacyXlsImportDiagnostic[] errors = workbook.Diagnostics
                .Where(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error)
                .ToArray();
            if (errors.Length > 0) {
                throw new InvalidDataException("Legacy XLS import failed: " + FormatLegacyXlsDiagnostics(errors));
            }

            return ProjectLoadedLegacyXlsWorkbook(workbook, filePath, readOnly);
        }

        private static string FormatLegacyXlsDiagnostics(IEnumerable<LegacyXlsImportDiagnostic> diagnostics) {
            const int maxDiagnostics = 6;
            LegacyXlsImportDiagnostic[] selected = diagnostics.Take(maxDiagnostics + 1).ToArray();
            string message = string.Join("; ", selected.Take(maxDiagnostics).Select(diagnostic => diagnostic.ToString()));
            if (selected.Length > maxDiagnostics) {
                message += $"; +{selected.Length - maxDiagnostics} more";
            }

            return message;
        }

        private static ExcelDocument LoadXlsbFromNormalFlow(
            byte[] bytes,
            bool readOnly,
            bool saveOnDispose,
            string? filePath,
            XlsbImportOptions? importOptions) {
            if (!readOnly && saveOnDispose) {
                throw new NotSupportedException("SaveOnDispose is not supported when loading XLSB workbooks. Save explicitly to an .xlsx path while native XLSB writing is unavailable.");
            }

            XlsbWorkbook workbook = XlsbWorkbookReader.Load(bytes, importOptions);
            ExcelDocument document = XlsbWorkbookProjector.ToExcelDocument(workbook);
            if (readOnly) {
                document = ReopenProjectedWorkbookReadOnly(document);
            }
            document.MarkLoadedFromXlsb(filePath, workbook);
            return document;
        }

        /// <summary>
        /// Loads an existing Excel document.
        /// </summary>
        /// <param name="filePath">Path to the file.</param>
        /// <param name="options">Access, persistence, and low-level package options.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Load(string filePath, ExcelLoadOptions? options = null) {
            if (filePath == null) {
                throw new ArgumentNullException(nameof(filePath));
            }

            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            byte[] bytes;
            using (var source = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete)) {
                bytes = ReadAllBytes(source, resolved.PackageSecurity);
            }
            return LoadFromByteArray(bytes, resolved, filePath);
        }

        /// <summary>
        /// Loads a password-encrypted Office Open XML workbook or legacy binary `.xls` workbook.
        /// </summary>
        /// <param name="filePath">Path to the encrypted workbook.</param>
        /// <param name="password">Password used to decrypt the workbook package.</param>
        /// <param name="options">Access and load options. SaveOnDispose is not supported for encrypted sources.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument LoadEncrypted(string filePath, string password, ExcelLoadOptions? options = null) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            byte[] encryptedBytes;
            using (var source = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete)) {
                encryptedBytes = ReadAllBytes(source, resolved.PackageSecurity);
            }
            if (resolved.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(encryptedBytes, resolved.PackageSecurity);
            }
            if (ExcelDocumentLoadRouting.IsEncryptedLegacyXls(encryptedBytes, filePath)) {
                return LoadEncryptedLegacyXls(encryptedBytes, password, resolved.AccessMode == DocumentAccessMode.ReadOnly);
            }

            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return LoadFromByteArray(packageBytes, resolved, filePath: null);
        }

        /// <summary>
        /// Loads an existing Excel document from the provided stream.
        /// </summary>
        /// <param name="stream">Input stream containing the workbook package. Editable writable seekable sources become the associated destination; other sources remain detached.</param>
        /// <param name="options">Access, persistence, and low-level package options.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Load(Stream stream, ExcelLoadOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "workbook");
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));

            var bytes = ReadAllBytes(stream, resolved.PackageSecurity);
            return LoadFromByteArray(
                bytes,
                resolved,
                filePath: null,
                originalStream: stream,
                leaveOriginalStreamOpen: true);
        }

        /// <summary>
        /// Loads a password-encrypted Office Open XML workbook or legacy binary workbook from a stream.
        /// </summary>
        /// <param name="stream">Input stream containing the encrypted workbook.</param>
        /// <param name="password">Password used to decrypt the workbook package.</param>
        /// <param name="options">Access and load options. SaveOnDispose is not supported for encrypted sources.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument LoadEncrypted(Stream stream, string password, ExcelLoadOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);

            var encryptedBytes = ReadAllBytes(stream, resolved.PackageSecurity);
            if (resolved.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(encryptedBytes, resolved.PackageSecurity);
            }
            if (ExcelDocumentLoadRouting.IsEncryptedLegacyXls(encryptedBytes, filePath: null)) {
                return LoadEncryptedLegacyXls(encryptedBytes, password, resolved.AccessMode == DocumentAccessMode.ReadOnly);
            }

            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return LoadFromByteArray(packageBytes, resolved, filePath: null);
        }

        /// <summary>
        /// Validates the current spreadsheet with Open XML validator and returns error messages (if any).
        /// Useful for troubleshooting "Repaired Records" issues in Excel.
        /// </summary>
        public System.Collections.Generic.IReadOnlyList<string> ValidateOpenXml() {
            var list = new System.Collections.Generic.List<string>();
            if (_spreadSheetDocument == null) return list;
            // Ensure worksheet element order prior to validation so schema checks reflect final layout
            try {
                foreach (var sheet in Sheets) {
                    sheet.EnsureWorksheetElementOrder();
                }
            } catch { }
            var validator = new OpenXmlValidator();
            foreach (var error in validator.Validate(_spreadSheetDocument)) {
                list.Add($"{error.ErrorType}: {error.Description} at {error.Path}");
            }
            return list;
        }

        private static System.Collections.Generic.IReadOnlyList<string> ValidateOpenXml(byte[] packageBytes) {
            var list = new System.Collections.Generic.List<string>();
            if (packageBytes == null || packageBytes.Length == 0) return list;

            using var stream = new MemoryStream(packageBytes, writable: false);
            using var document = SpreadsheetDocument.Open(stream, false);
            var validator = new OpenXmlValidator();
            foreach (var error in validator.Validate(document)) {
                list.Add($"{error.ErrorType}: {error.Description} at {error.Path}");
            }
            return list;
        }

        /// <summary>
        /// Asynchronously loads an Excel document from the specified path.
        /// </summary>
        /// <param name="filePath">Path to the Excel file.</param>
        /// <param name="options">Access, persistence, and low-level package options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        /// <exception cref="FileNotFoundException">Thrown when the file does not exist.</exception>
        public static async Task<ExcelDocument> LoadAsync(string filePath, ExcelLoadOptions? options = null, CancellationToken cancellationToken = default) {
            if (filePath == null) {
                throw new ArgumentNullException(nameof(filePath));
            }
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            var bytes = await ReadAllBytesCompatAsync(filePath, cancellationToken, resolved.PackageSecurity)
                .ConfigureAwait(false);
            return LoadFromByteArray(bytes, resolved, filePath);
        }

        /// <summary>
        /// Asynchronously loads a password-encrypted Office Open XML workbook or legacy binary `.xls` workbook.
        /// </summary>
        /// <param name="filePath">Path to the encrypted workbook.</param>
        /// <param name="password">Password used to decrypt the workbook package.</param>
        /// <param name="options">Access and load options. SaveOnDispose is not supported for encrypted sources.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static async Task<ExcelDocument> LoadEncryptedAsync(string filePath, string password, ExcelLoadOptions? options = null, CancellationToken cancellationToken = default) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            var encryptedBytes = await ReadAllBytesCompatAsync(filePath, cancellationToken, resolved.PackageSecurity)
                .ConfigureAwait(false);
            if (resolved.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(encryptedBytes, resolved.PackageSecurity);
            }
            if (ExcelDocumentLoadRouting.IsEncryptedLegacyXls(encryptedBytes, filePath)) {
                return LoadEncryptedLegacyXls(encryptedBytes, password, resolved.AccessMode == DocumentAccessMode.ReadOnly);
            }

            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return LoadFromByteArray(packageBytes, resolved, filePath: null);
        }

        /// <summary>
        /// Asynchronously loads an Excel document from the provided stream.
        /// </summary>
        /// <param name="stream">Input stream containing the workbook package. Editable writable seekable sources become the associated destination; other sources remain detached.</param>
        /// <param name="options">Access, persistence, and low-level package options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static async Task<ExcelDocument> LoadAsync(Stream stream, ExcelLoadOptions? options = null, CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "workbook");
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));

            var bytes = await ReadAllBytesAsync(stream, cancellationToken, resolved.PackageSecurity)
                .ConfigureAwait(false);
            return LoadFromByteArray(
                bytes,
                resolved,
                filePath: null,
                originalStream: stream,
                leaveOriginalStreamOpen: true);
        }

        /// <summary>
        /// Asynchronously loads a password-encrypted Office Open XML workbook or legacy binary workbook from a stream.
        /// </summary>
        /// <param name="stream">Input stream containing the encrypted workbook.</param>
        /// <param name="password">Password used to decrypt the workbook package.</param>
        /// <param name="options">Access and load options. SaveOnDispose is not supported for encrypted sources.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static async Task<ExcelDocument> LoadEncryptedAsync(Stream stream, string password, ExcelLoadOptions? options = null, CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            EnsureEncryptedLoadUsesExplicitPersistence(resolved);

            var encryptedBytes = await ReadAllBytesAsync(stream, cancellationToken, resolved.PackageSecurity)
                .ConfigureAwait(false);
            if (resolved.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(encryptedBytes, resolved.PackageSecurity);
            }
            if (ExcelDocumentLoadRouting.IsEncryptedLegacyXls(encryptedBytes, filePath: null)) {
                return LoadEncryptedLegacyXls(encryptedBytes, password, resolved.AccessMode == DocumentAccessMode.ReadOnly);
            }

            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return LoadFromByteArray(packageBytes, resolved, filePath: null);
        }

        private static ExcelDocument LoadEncryptedLegacyXls(byte[] bytes, string password, bool readOnly) {
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(bytes, new LegacyXlsImportOptions {
                Password = password,
                ReportUnsupportedContent = true
            });
            return ProjectLoadedLegacyXlsWorkbook(workbook, sourcePath: null, readOnly);
        }

        private static void EnsureEncryptedLoadUsesExplicitPersistence(ExcelLoadOptions options) {
            if (options.PersistenceMode != DocumentPersistenceMode.Explicit) {
                throw new NotSupportedException("SaveOnDispose is not supported for encrypted Excel sources. Use SaveEncrypted to persist encrypted changes.");
            }
        }
    }
}
