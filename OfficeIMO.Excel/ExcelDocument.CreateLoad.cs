using OfficeIMO.Core.Internal;
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
            if (resolved.Template != null || resolved.TemplateStream != null) return CreateFromUnifiedOptions(null, null, resolved);
            if (resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose) {
                throw new ArgumentException("SaveOnDispose requires an associated file path or writable stream.", nameof(options));
            }

            var packageStream = new MemoryStream(StreamBufferSize);
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(packageStream, resolved.DocumentType.ToOpenXml(), autoSave: false);
            return CreateNewDocument(
                spreadSheetDocument,
                filePath: null,
                packageStream,
                sourceStream: null,
                resolved.PersistenceMode,
                copyPackageToSourceOnDispose: false,
                leaveSourceStreamOpen: true, firstSheetName: resolved.FirstSheetName);
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
            if (resolved.Template != null || resolved.TemplateStream != null) return CreateFromUnifiedOptions(filePath, null, resolved);
            bool saveOnDispose = resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose;
            if (saveOnDispose && string.IsNullOrEmpty(filePath)) {
                throw new ArgumentException("SaveOnDispose requires an associated file path or writable stream.", nameof(filePath));
            }

            Stream packageStream = saveOnDispose
                ? new NonDisposingMemoryStream(StreamBufferSize)
                : new MemoryStream(StreamBufferSize);
            SpreadsheetDocumentType documentType = ResolveSpreadsheetDocumentType(filePath) ?? resolved.DocumentType.ToOpenXml();
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(packageStream, documentType, autoSave: false);
            return CreateNewDocument(
                spreadSheetDocument,
                filePath,
                packageStream,
                sourceStream: null,
                resolved.PersistenceMode,
                copyPackageToSourceOnDispose: false,
                leaveSourceStreamOpen: true,
                copyPackageToFilePathOnDispose: saveOnDispose, firstSheetName: resolved.FirstSheetName);
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
            if (resolved.Template != null || resolved.TemplateStream != null) return CreateFromUnifiedOptions(null, stream, resolved);
            bool saveOnDispose = resolved.PersistenceMode == DocumentPersistenceMode.SaveOnDispose;

            Stream packageStream = saveOnDispose
                ? new NonDisposingMemoryStream(StreamBufferSize)
                : new MemoryStream(StreamBufferSize);

            var spreadSheetDocument = SpreadsheetDocument.Create(packageStream, resolved.DocumentType.ToOpenXml(), false);
            return CreateNewDocument(spreadSheetDocument, filePath: null, packageStream, stream, resolved.PersistenceMode, saveOnDispose, leaveSourceStreamOpen: true, firstSheetName: resolved.FirstSheetName);
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
                ChangeSpreadsheetDocumentType(target.Value);
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
            Stream? ownedOpenStream = null,
            string? firstSheetName = null) {
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
            document.EnsureWorkbookThemeAndStyles();

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
            document.BuiltinDocumentProperties = new ExcelBuiltinDocumentProperties(document);
            document.ApplicationProperties = new ExcelApplicationProperties(document);
            document.CustomDocumentProperties.SetChangeHandler(document.MarkCustomDocumentPropertiesChanged);
            document.LoadCustomDocumentProperties();

            if (firstSheetName != null) document.AddWorksheet(firstSheetName);
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

            document.BuiltinDocumentProperties = new ExcelBuiltinDocumentProperties(document);
            document.ApplicationProperties = new ExcelApplicationProperties(document);
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
            bool leaveOriginalStreamOpen = true,
            bool captureSourcePackageSecurity = true) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));
            if (options == null) throw new ArgumentNullException(nameof(options));
            OfficeDocumentLifecycle.Validate(options.AccessMode, options.PersistenceMode, "workbook");
            long? inputLimit = ResolveInputLimit(options);
            if (inputLimit.HasValue && bytes.LongLength > inputLimit.Value) {
                throw new InvalidDataException(
                    $"Workbook input contains {bytes.LongLength} bytes, exceeding MaxInputBytes ({inputLimit.Value}).");
            }
            if (options.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(bytes, options.PackageSecurity);
            }

            bool readOnly = options.AccessMode == DocumentAccessMode.ReadOnly;
            bool saveOnDispose = options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose;
            Stream? associatedStream = OfficeDocumentLifecycle.ResolveAssociatedDestination(
                originalStream,
                options.AccessMode);

            ExcelFileFormat detectedFormat = ExcelDocumentLoadRouting.DetectFormat(bytes, filePath);
            if (!options.AllowBinaryProjection && detectedFormat is ExcelFileFormat.Xls or ExcelFileFormat.Xlsb) {
                throw new NotSupportedException("This load profile accepts XML workbook packages only; XLS and XLSB projection is disabled.");
            }
            if (detectedFormat == ExcelFileFormat.Xls) {
                ExcelDocument document = LoadLegacyXlsFromNormalFlow(
                    bytes, readOnly, saveOnDispose, filePath, CreateLegacyImportOptions(options));
                return captureSourcePackageSecurity ? WithSourcePackageInspection(document, bytes, options) : document;
            }

            if (detectedFormat == ExcelFileFormat.Xlsb) {
                ExcelDocument document = LoadXlsbFromNormalFlow(bytes, readOnly, saveOnDispose, filePath, options.XlsbImportOptions);
                return captureSourcePackageSecurity ? WithSourcePackageInspection(document, bytes, options) : document;
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

                bool normalizedContentTypes = Utilities.ExcelPackageUtilities.NormalizeContentTypes(
                    normalizedStream,
                    leaveOpen: true,
                    maxCharactersInPart: effectiveOpenSettings.MaxCharactersInPart);
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
                document._sourceMaxCharactersInPart = effectiveOpenSettings.MaxCharactersInPart;
                document._openXmlOriginalPackageBytes = OfficeCompatibilitySourceCarrier.ContainsPackageCarrier(bytes)
                    ? (byte[])bytes.Clone()
                    : null;
                return captureSourcePackageSecurity ? WithSourcePackageInspection(document, bytes, options) : document;
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

        private static byte[] ReadAllBytes(Stream stream, ExcelLoadOptions options) {
            long? inputLimit = ResolveInputLimit(options);
            return options.PackageSecurity != null
                && inputLimit == options.PackageSecurity.MaxPackageBytes
                ? OfficePackageSecurityInspector.ReadBounded(stream, options.PackageSecurity)
                : OfficeStreamReader.ReadAllBytes(stream, inputLimit);
        }

        private static async Task<byte[]> ReadAllBytesAsync(Stream stream, CancellationToken cancellationToken,
            ExcelLoadOptions options) {
            long? inputLimit = ResolveInputLimit(options);
            return options.PackageSecurity != null
                && inputLimit == options.PackageSecurity.MaxPackageBytes
                ? await OfficePackageSecurityInspector.ReadBoundedAsync(stream, options.PackageSecurity, cancellationToken)
                    .ConfigureAwait(false)
                : await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, inputLimit)
                    .ConfigureAwait(false);
        }

        private static long? ResolveInputLimit(ExcelLoadOptions options) {
            long? configured = options.MaxInputBytes;
            if (configured.HasValue && configured.Value < 1) {
                throw new ArgumentOutOfRangeException(nameof(options.MaxInputBytes));
            }
            if (options.PackageSecurity == null) return configured;
            long packageLimit = options.PackageSecurity.MaxPackageBytes;
            if (packageLimit < 1) {
                throw new ArgumentOutOfRangeException(nameof(OfficePackageSecurityOptions.MaxPackageBytes));
            }
            return configured.HasValue ? Math.Min(configured.Value, packageLimit) : packageLimit;
        }

        private static LegacyXlsImportOptions CreateLegacyImportOptions(ExcelLoadOptions options) {
            long? limit = ResolveInputLimit(options);
            LegacyXlsImportOptions legacy = options.CreateLegacyXlsOptions();
            if (limit.HasValue) legacy.MaxInputBytes = checked((int)Math.Min(legacy.MaxInputBytes, limit.Value));
            return legacy;
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

        internal static ExcelDocument LoadLegacyXlsFromNormalFlow(
            byte[] bytes,
            bool readOnly,
            bool saveOnDispose,
            string? filePath,
            LegacyXlsImportOptions? importOptions = null,
            Action<LegacyXlsWorkbook>? beforeProjection = null) {
            if (!readOnly && saveOnDispose) {
                throw new NotSupportedException("SaveOnDispose is not supported for imported XLS workbooks. Use an explicit Save operation.");
            }

            LegacyXlsImportOptions effectiveOptions = importOptions ?? new LegacyXlsImportOptions();
            effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(bytes, effectiveOptions);
            effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
            LegacyXlsImportDiagnostic[] errors = workbook.Diagnostics
                .Where(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error)
                .ToArray();
            if (errors.Length > 0) {
                throw new InvalidDataException("Legacy XLS import failed: " + FormatLegacyXlsDiagnostics(errors));
            }
            beforeProjection?.Invoke(workbook);

            return ProjectLoadedLegacyXlsWorkbook(
                workbook,
                filePath,
                readOnly,
                effectiveOptions.CancellationToken);
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
                throw new NotSupportedException("SaveOnDispose is not supported for imported XLSB workbooks. Use an explicit Save operation.");
            }

            XlsbImportOptions effectiveOptions = importOptions ?? new XlsbImportOptions();
            XlsbWorkbook workbook = XlsbWorkbookReader.Load(bytes, effectiveOptions);
            ExcelDocument document = XlsbWorkbookProjector.ToExcelDocument(workbook,
                effectiveOptions.CancellationToken, effectiveOptions.MaxCells);
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
            if (options?.Password is string password) return AssociateEncryptedWorkbook(LoadEncrypted(filePath, password, options), filePath, password, options: options);
            if (options?.FileBacked == true) return OpenFileBacked(filePath, options);
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
                bytes = ReadAllBytes(source, resolved);
            }
            return LoadFromByteArray(bytes, resolved, filePath);
        }

        /// <summary>
        /// Loads a password-encrypted Office Open XML workbook or legacy binary `.xls` workbook.
        /// </summary>
        /// <param name="filePath">Path to the encrypted workbook.</param>
        /// <param name="password">Password used to decrypt the workbook package.</param>
        /// <param name="options">Access and load options for the decrypted package.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        internal static ExcelDocument LoadEncrypted(string filePath, string password, ExcelLoadOptions? options = null) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "document");
            resolved = resolved.Snapshot();
            resolved.PersistenceMode = DocumentPersistenceMode.Explicit;
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            byte[] encryptedBytes;
            using (var source = new FileStream(filePath, FileMode.Open, FileAccess.Read,
                FileShare.ReadWrite | FileShare.Delete)) {
                encryptedBytes = ReadAllBytes(source, resolved);
            }
            if (resolved.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(encryptedBytes, resolved.PackageSecurity);
            }
            if (ExcelDocumentLoadRouting.IsEncryptedLegacyXls(encryptedBytes, filePath)) {
                return WithSourcePackageInspection(LoadEncryptedLegacyXls(encryptedBytes, password, resolved.AccessMode == DocumentAccessMode.ReadOnly, resolved), encryptedBytes, resolved, "encrypted-source-container");
            }

            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return WithSourcePackageInspection(LoadFromByteArray(packageBytes, resolved, filePath: null, captureSourcePackageSecurity: false), encryptedBytes, resolved, "encrypted-source-container");
        }

        /// <summary>
        /// Loads an existing Excel document from the provided stream.
        /// </summary>
        /// <param name="stream">Input stream containing the workbook package. Editable writable seekable sources become the associated destination; other sources remain detached.</param>
        /// <param name="options">Access, persistence, and low-level package options.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Load(Stream stream, ExcelLoadOptions? options = null) {
            if (options?.Password is string password) return AssociateEncryptedWorkbook(LoadEncrypted(stream, password, options), null, password, options.AccessMode == DocumentAccessMode.ReadOnly ? null : stream, options);
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "workbook");
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));

            var bytes = ReadAllBytes(stream, resolved);
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
        /// <param name="options">Access and load options for the decrypted package.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        internal static ExcelDocument LoadEncrypted(Stream stream, string password, ExcelLoadOptions? options = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "document");
            resolved = resolved.Snapshot();
            resolved.PersistenceMode = DocumentPersistenceMode.Explicit;

            var encryptedBytes = ReadAllBytes(stream, resolved);
            if (resolved.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(encryptedBytes, resolved.PackageSecurity);
            }
            if (ExcelDocumentLoadRouting.IsEncryptedLegacyXls(encryptedBytes, filePath: null)) {
                return WithSourcePackageInspection(LoadEncryptedLegacyXls(encryptedBytes, password, resolved.AccessMode == DocumentAccessMode.ReadOnly, resolved), encryptedBytes, resolved, "encrypted-source-container");
            }

            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return WithSourcePackageInspection(LoadFromByteArray(packageBytes, resolved, filePath: null, captureSourcePackageSecurity: false), encryptedBytes, resolved, "encrypted-source-container");
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

        /// <summary>
        /// Asynchronously loads an Excel document from the specified path.
        /// </summary>
        /// <param name="filePath">Path to the Excel file.</param>
        /// <param name="options">Access, persistence, and low-level package options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        /// <exception cref="FileNotFoundException">Thrown when the file does not exist.</exception>
        public static async Task<ExcelDocument> LoadAsync(string filePath, ExcelLoadOptions? options = null, CancellationToken cancellationToken = default) {
            options = (options ?? new ExcelLoadOptions()).Snapshot();
            options.XlsbImportOptions.CancellationToken = cancellationToken;
            cancellationToken.ThrowIfCancellationRequested();
            if (options?.Password is string password) return AssociateEncryptedWorkbook(await LoadEncryptedAsync(filePath, password, options, cancellationToken).ConfigureAwait(false), filePath, password, options: options);
            if (options?.FileBacked == true) return await Task.Run(() => OpenFileBacked(filePath, options, cancellationToken), cancellationToken).ConfigureAwait(false);
            if (filePath == null) {
                throw new ArgumentNullException(nameof(filePath));
            }
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            var bytes = await ReadAllBytesCompatAsync(filePath, cancellationToken, resolved)
                .ConfigureAwait(false);
            return LoadFromByteArray(bytes, resolved, filePath);
        }

        /// <summary>
        /// Asynchronously loads a password-encrypted Office Open XML workbook or legacy binary `.xls` workbook.
        /// </summary>
        /// <param name="filePath">Path to the encrypted workbook.</param>
        /// <param name="password">Password used to decrypt the workbook package.</param>
        /// <param name="options">Access and load options for the decrypted package.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        internal static async Task<ExcelDocument> LoadEncryptedAsync(string filePath, string password, ExcelLoadOptions? options = null, CancellationToken cancellationToken = default) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "document");
            resolved = resolved.Snapshot();
            resolved.PersistenceMode = DocumentPersistenceMode.Explicit;
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            var encryptedBytes = await ReadAllBytesCompatAsync(filePath, cancellationToken, resolved)
                .ConfigureAwait(false);
            if (resolved.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(encryptedBytes, resolved.PackageSecurity);
            }
            if (ExcelDocumentLoadRouting.IsEncryptedLegacyXls(encryptedBytes, filePath)) {
                return WithSourcePackageInspection(LoadEncryptedLegacyXls(encryptedBytes, password, resolved.AccessMode == DocumentAccessMode.ReadOnly, resolved), encryptedBytes, resolved, "encrypted-source-container");
            }

            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return WithSourcePackageInspection(LoadFromByteArray(packageBytes, resolved, filePath: null, captureSourcePackageSecurity: false), encryptedBytes, resolved, "encrypted-source-container");
        }

        /// <summary>
        /// Asynchronously loads an Excel document from the provided stream.
        /// </summary>
        /// <param name="stream">Input stream containing the workbook package. Editable writable seekable sources become the associated destination; other sources remain detached.</param>
        /// <param name="options">Access, persistence, and low-level package options.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static async Task<ExcelDocument> LoadAsync(Stream stream, ExcelLoadOptions? options = null, CancellationToken cancellationToken = default) {
            options = (options ?? new ExcelLoadOptions()).Snapshot();
            options.XlsbImportOptions.CancellationToken = cancellationToken;
            cancellationToken.ThrowIfCancellationRequested();
            if (options?.Password is string password) return AssociateEncryptedWorkbook(await LoadEncryptedAsync(stream, password, options, cancellationToken).ConfigureAwait(false), null, password, options.AccessMode == DocumentAccessMode.ReadOnly ? null : stream, options);
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "workbook");
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));

            var bytes = await ReadAllBytesAsync(stream, cancellationToken, resolved)
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
        /// <param name="options">Access and load options for the decrypted package.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        internal static async Task<ExcelDocument> LoadEncryptedAsync(Stream stream, string password, ExcelLoadOptions? options = null, CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "document");
            resolved = resolved.Snapshot();
            resolved.PersistenceMode = DocumentPersistenceMode.Explicit;

            var encryptedBytes = await ReadAllBytesAsync(stream, cancellationToken, resolved)
                .ConfigureAwait(false);
            if (resolved.PackageSecurity != null) {
                OfficePackageSecurityInspector.Validate(encryptedBytes, resolved.PackageSecurity);
            }
            if (ExcelDocumentLoadRouting.IsEncryptedLegacyXls(encryptedBytes, filePath: null)) {
                return WithSourcePackageInspection(LoadEncryptedLegacyXls(encryptedBytes, password, resolved.AccessMode == DocumentAccessMode.ReadOnly, resolved), encryptedBytes, resolved, "encrypted-source-container");
            }

            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return WithSourcePackageInspection(LoadFromByteArray(packageBytes, resolved, filePath: null, captureSourcePackageSecurity: false), encryptedBytes, resolved, "encrypted-source-container");
        }

        private static ExcelDocument LoadEncryptedLegacyXls(byte[] bytes, string password, bool readOnly, ExcelLoadOptions options) {
            if (!options.AllowBinaryProjection) throw new NotSupportedException("This load profile accepts XML workbook packages only; encrypted XLS projection is disabled.");
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(bytes, new LegacyXlsImportOptions {
                Password = password,
                ReportUnsupportedContent = true
            });
            if (!workbook.WasEncryptedSource) {
                throw new InvalidDataException("LoadEncrypted requires a password-encrypted legacy XLS workbook.");
            }
            return ProjectLoadedLegacyXlsWorkbook(workbook, sourcePath: null, readOnly);
        }

    }
}
