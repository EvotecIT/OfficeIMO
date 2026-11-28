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
using System.Xml;
using System.Diagnostics;

namespace OfficeIMO.Excel {
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {

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

            document.Properties = new ExcelBuiltinDocumentProperties(document);
            document.ApplicationProperties = new ExcelApplicationProperties(document);
            document.CustomProperties.SetChangeHandler(document.MarkCustomDocumentPropertiesChanged);
            document.LoadCustomDocumentProperties();
            ExcelChartAxisIdGenerator.ReadChartRelationships(document._spreadSheetDocument);
            return document;
        }

        /// <summary>
        /// Creates a new workbook over a package OfficeIMO writes itself; the SDK opens over it when something needs the model.
        /// </summary>
        private static ExcelDocument CreateOwnedDocument(ExcelCreateOptions options, ExcelDocumentType type, string? filePath, Stream? sourceStream) {
            string mainContentType = CreatedMainContentType(type) ?? throw new ArgumentOutOfRangeException(nameof(options), type, "Unknown workbook type.");
            byte[] bytes = CreateWorkbookShell(mainContentType, options.Theme);
            var deferred = new ExcelDeferredWorkbook(CreatedWorkbookPartName, new List<ExcelDeferredSheet>(), date1904: false);
            // A workbook saved on dispose writes itself to its file or stream then, as any dirty workbook does.
            ExcelDocument document = CreateDeferredDocument(deferred, bytes, CreateOpenSettings(null), readOnly: false, filePath, sourceStream,
                leaveSourceStreamOpen: true, options.PersistenceMode);
            document._deferredCreated = true;
            document._deferredPackageWrittenHere = true;
            // Its destination receives the workbook on dispose even when a copy was saved elsewhere meanwhile.
            bool saveOnDispose = options.PersistenceMode == DocumentPersistenceMode.SaveOnDispose;
            document._copyPackageToSourceOnDispose = saveOnDispose && sourceStream != null;
            document._copyPackageToFilePathOnDispose = saveOnDispose && filePath != null;
            document._creationTheme = options.Theme ?? OfficeTheme.Default;
            document._defaultFontTheme = options.Theme;
            // A new workbook is saved even before anything changes, as one created through the model is.
            document._packageDirty = true;
            document._unchangedPackageBytes = null;
            // Its package is OfficeIMO's own, so the fast writers may rewrite it as they do a workbook the model creates.
            document._packageContentTypesKnownNormalized = false;
            if (options.FirstSheetName != null) document.AddSheet(options.FirstSheetName);
            return document;
        }

        /// <summary>Creates a document over a loaded package whose SDK opens on first use; see <see cref="OpenDeferredSdk"/>.</summary>
        private static ExcelDocument CreateDeferredDocument(
            ExcelDeferredWorkbook deferred,
            byte[] packageBytes,
            OpenSettings openSettings,
            bool readOnly,
            string? filePath,
            Stream? sourceStream,
            bool leaveSourceStreamOpen,
            DocumentPersistenceMode persistenceMode) {
            var document = new ExcelDocument {
                FilePath = filePath,
                _deferredWorkbook = deferred,
                _deferredPackageBytes = packageBytes,
                _deferredOpenSettings = openSettings,
                _deferredReadOnly = readOnly,
                _sourceStream = sourceStream,
                _leaveSourceStreamOpen = leaveSourceStreamOpen,
                _packageContentTypesKnownNormalized = true,
                _simplePackageContentKnown = false,
                _requiresSavePreflight = false,
                _packageDirty = false,
                _packagePropertiesDirty = false,
                _unchangedPackageBytes = packageBytes,
                _persistenceMode = persistenceMode,
                _sourceMaxCharactersInPart = openSettings.MaxCharactersInPart,
                _readSnapshotAllowed = true,
            };

            document.Properties = new ExcelBuiltinDocumentProperties(document);
            document.ApplicationProperties = new ExcelApplicationProperties(document);
            document.CustomProperties.SetChangeHandler(document.MarkCustomDocumentPropertiesChanged);
            // The load declined packages with custom properties, so there are none to read.
            document.CustomProperties.ReplaceWith(new Dictionary<string, ExcelCustomProperty>(StringComparer.OrdinalIgnoreCase));
            return document;
        }

        private static ExcelDocument CreateDeferredDocument(ExcelDeferredWorkbook deferred, byte[] packageBytes, bool readOnly) =>
            CreateDeferredDocument(deferred, packageBytes, CreateOpenSettings(null), readOnly,
                filePath: null, sourceStream: null, leaveSourceStreamOpen: true, DocumentPersistenceMode.Explicit);

        private static WorkbookPart GetWorkbookPartOrThrow(SpreadsheetDocument document) {
            ArgumentNullException.ThrowIfNull(document);

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
            ArgumentNullException.ThrowIfNull(bytes);
            ArgumentNullException.ThrowIfNull(options);
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

            ExcelFileFormat detectedFormat = ExcelDocumentLoadRouting.DetectFormat(bytes, filePath,
                options.PackageSecurity?.MaxPartCount);
            if (!options.AllowBinaryProjection && detectedFormat is ExcelFileFormat.Xls or ExcelFileFormat.Xlsb) {
                throw new NotSupportedException("This load profile accepts XML workbook packages only; XLS and XLSB projection is disabled.");
            }
            if (detectedFormat == ExcelFileFormat.Xls) {
                ExcelDocument document = LoadLegacyXlsFromNormalFlow(
                    bytes, readOnly, saveOnDispose, filePath, CreateLegacyImportOptions(options),
                    earlyBiffSecurityPolicy: options.PackageSecurity);
                document.RememberEarlyBiffSource(originalStream);
                return captureSourcePackageSecurity ? WithSourcePackageInspection(document, bytes, options) : document;
            }

            if (detectedFormat == ExcelFileFormat.Xlsb) {
                ExcelDocument document = LoadXlsbFromNormalFlow(bytes, readOnly, saveOnDispose, filePath, options.XlsbImportOptions,
                    callerValidatedPackage: options.PackageSecurity != null);
                return captureSourcePackageSecurity ? WithSourcePackageInspection(document, bytes, options) : document;
            }

            var effectiveOpenSettings = CreateOpenSettings(options.OpenSettings);
            bool shouldCopyBack = saveOnDispose && associatedStream != null;
            bool shouldCopyBackToFilePath = !shouldCopyBack && !string.IsNullOrEmpty(filePath) && saveOnDispose;
            bool shouldRetainPackageStream = shouldCopyBack || shouldCopyBackToFilePath;

            Stream? normalizedStream = null;

            try {
                byte[] unchangedPackageBytes;
                if (!shouldRetainPackageStream && Utilities.ExcelPackageUtilities.TryNormalizeContentTypes(
                        bytes, effectiveOpenSettings.MaxCharactersInPart, out byte[] normalizedBytes, out _)) {
                    ExcelDeferredWorkbook? deferred = options.PersistenceMode == DocumentPersistenceMode.Explicit
                        && !OfficeCompatibilitySourceCarrier.ContainsPackageCarrier(bytes)
                        && IsPlainOpenSettings(effectiveOpenSettings)
                        ? TryReadDeferredWorkbook(normalizedBytes)
                        : null;
                    if (deferred != null) {
                        ExcelDocument deferredDocument = CreateDeferredDocument(deferred, normalizedBytes, effectiveOpenSettings, readOnly,
                            filePath, associatedStream, leaveOriginalStreamOpen, options.PersistenceMode);
                        return captureSourcePackageSecurity ? WithSourcePackageInspection(deferredDocument, bytes, options) : deferredDocument;
                    }
                    normalizedStream = OfficePackageLoadSource.CreatePackageStream(normalizedBytes, DocumentAccessMode.ReadWrite);
                    unchangedPackageBytes = normalizedBytes;
                } else {
                    // A package copied back on dispose is read and written over the loaded array, which it copies only
                    // when it first writes. Like NonDisposingMemoryStream, the stream keeps its bytes after the SDK
                    // closes it, until DisposeStream releases it.
                    Stream packageStream = shouldRetainPackageStream
                        ? new OfficeCopyOnWriteStream(bytes, survivesDispose: true)
                        : new MemoryStream(bytes.Length + StreamBufferSize);
                    normalizedStream = packageStream;
                    if (!shouldRetainPackageStream) {
                        packageStream.Write(bytes, 0, bytes.Length);
                        packageStream.Position = 0;
                    }

                    bool normalizedContentTypes = Utilities.ExcelPackageUtilities.NormalizeContentTypes(
                        packageStream,
                        leaveOpen: true,
                        maxCharactersInPart: effectiveOpenSettings.MaxCharactersInPart);
                    packageStream.Position = 0;
                    unchangedPackageBytes = normalizedContentTypes ? OfficePackageLoadSource.ReadPackageBytes(packageStream) : bytes;
                }

                var memDoc = readOnly
                    ? OfficeIMO.OpenXml.Internal.OfficeOpenXmlReadOnlyPackage.OpenSpreadsheet(normalizedStream, effectiveOpenSettings)
                    : SpreadsheetDocument.Open(normalizedStream, true, effectiveOpenSettings);
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
                // A retained package must still write its final archive for the copy back, so it is never abandoned.
                document._packageCopyOnWriteStream = shouldRetainPackageStream ? null : normalizedStream as OfficeCopyOnWriteStream;
                document._packageOnCopyOnWriteStream = document._packageCopyOnWriteStream != null ? memDoc : null;
                document._sourceMaxCharactersInPart = effectiveOpenSettings.MaxCharactersInPart;
                // The streaming readers apply neither setting, so those loads stay on the SDK route.
                document._readSnapshotAllowed = effectiveOpenSettings.MaxCharactersInPart == 0
                    && effectiveOpenSettings.MarkupCompatibilityProcessSettings.ProcessMode == MarkupCompatibilityProcessMode.NoProcess;
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
            var probeLimit = Legacy.EarlyBiffWorkbook.ProbeInputLimitFor(options.MaxLegacyInputBytes);
            return options.PackageSecurity != null
                && inputLimit == options.PackageSecurity.MaxPackageBytes
                ? OfficePackageSecurityInspector.ReadBounded(stream, options.PackageSecurity, options.LoadCancellationToken,
                    Legacy.EarlyBiffWorkbook.ProbeLength, probeLimit)
                : OfficeStreamReader.ReadAllBytes(stream, options.LoadCancellationToken, inputLimit, Legacy.EarlyBiffWorkbook.ProbeLength, probeLimit);
        }

        private static async Task<byte[]> ReadAllBytesAsync(Stream stream, CancellationToken cancellationToken,
            ExcelLoadOptions options) {
            long? inputLimit = ResolveInputLimit(options);
            var probeLimit = Legacy.EarlyBiffWorkbook.ProbeInputLimitFor(options.MaxLegacyInputBytes);
            return options.PackageSecurity != null
                && inputLimit == options.PackageSecurity.MaxPackageBytes
                ? await OfficePackageSecurityInspector.ReadBoundedAsync(stream, options.PackageSecurity, cancellationToken, Legacy.EarlyBiffWorkbook.ProbeLength, probeLimit)
                    .ConfigureAwait(false)
                : await OfficeStreamReader.ReadAllBytesAsync(stream, cancellationToken, inputLimit, Legacy.EarlyBiffWorkbook.ProbeLength, probeLimit)
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
            legacy.CancellationToken = options.XlsbImportOptions.CancellationToken;
            return legacy;
        }

        private static void DisposeStream(Stream? stream) {
            if (stream == null) {
                return;
            }

            if (stream is NonDisposingMemoryStream ndms) {
                ndms.DisposeUnderlying();
            } else if (stream is OfficeCopyOnWriteStream copyOnWrite) {
                copyOnWrite.Release();
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
            Action<LegacyXlsWorkbook>? beforeProjection = null,
            OfficePackageSecurityOptions? earlyBiffSecurityPolicy = null) {
            if (!readOnly && saveOnDispose) {
                throw new NotSupportedException("SaveOnDispose is not supported for imported XLS workbooks. Use an explicit Save operation.");
            }

            LegacyXlsImportOptions effectiveOptions = importOptions ?? new LegacyXlsImportOptions();
            effectiveOptions.CancellationToken.ThrowIfCancellationRequested();
            if (Legacy.EarlyBiffWorkbook.Identify(bytes) != null) {
                return LoadEarlyBiff(bytes, readOnly, filePath, effectiveOptions, earlyBiffSecurityPolicy);
            }

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
            XlsbImportOptions? importOptions,
            bool callerValidatedPackage) {
            if (!readOnly && saveOnDispose) {
                throw new NotSupportedException("SaveOnDispose is not supported for imported XLSB workbooks. Use an explicit Save operation.");
            }

            if (!callerValidatedPackage) {
                OfficePackageSecurityInspector.Validate(bytes, OfficePackageSecurityOptions.SecureDefaults);
            }

            XlsbImportOptions effectiveOptions = importOptions ?? new XlsbImportOptions();
            XlsbWorkbook workbook = XlsbWorkbookReader.Load(bytes, effectiveOptions);
            ExcelDocument? owned = TryProjectXlsbOnOwnedPackage(workbook, readOnly,
                effectiveOptions.CancellationToken, effectiveOptions.MaxCells);
            if (owned != null) {
                try {
                    owned.MarkLoadedFromOwnedXlsb(filePath, workbook);
                    return owned;
                } catch {
                    owned.Dispose();
                    throw;
                }
            }
            ExcelDocument document = XlsbWorkbookProjector.ToExcelDocument(workbook,
                effectiveOptions.CancellationToken, effectiveOptions.MaxCells);
            try {
                XlsbDocumentPropertiesProjector.Apply(document, workbook);
            } catch {
                document.Dispose();
                throw;
            }
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
            if (options?.FileBacked == true) return OpenFileBacked(filePath, options, options.LoadCancellationToken);
            ArgumentNullException.ThrowIfNull(filePath);

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
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentNullException.ThrowIfNull(password);
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
                return WithSourcePackageInspection(LoadEncryptedLegacyXls(encryptedBytes, password, resolved.AccessMode == DocumentAccessMode.ReadOnly, resolved, resolved.LoadCancellationToken), encryptedBytes, resolved, "encrypted-source-container");
            }

            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password, resolved.LoadCancellationToken);
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
            ArgumentNullException.ThrowIfNull(stream);
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
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(password);
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
                return WithSourcePackageInspection(LoadEncryptedLegacyXls(encryptedBytes, password, resolved.AccessMode == DocumentAccessMode.ReadOnly, resolved, resolved.LoadCancellationToken), encryptedBytes, resolved, "encrypted-source-container");
            }

            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password, resolved.LoadCancellationToken);
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
            options.LoadCancellationToken = cancellationToken;
            cancellationToken.ThrowIfCancellationRequested();
            if (options?.Password is string password) return CompleteCancellableLoad(AssociateEncryptedWorkbook(await LoadEncryptedAsync(filePath, password, options, cancellationToken).ConfigureAwait(false), filePath, password, options: options), cancellationToken);
            if (options?.FileBacked == true) return CompleteCancellableLoad(await Task.Run(() => OpenFileBacked(filePath, options, cancellationToken), cancellationToken).ConfigureAwait(false), cancellationToken);
            ArgumentNullException.ThrowIfNull(filePath);
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            var bytes = await ReadAllBytesCompatAsync(filePath, cancellationToken, resolved)
                .ConfigureAwait(false);
            return CompleteCancellableLoad(LoadFromByteArray(bytes, resolved, filePath), cancellationToken);
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
            ArgumentNullException.ThrowIfNull(filePath);
            ArgumentNullException.ThrowIfNull(password);
            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "document");
            resolved = resolved.Snapshot();
            resolved.PersistenceMode = DocumentPersistenceMode.Explicit;
            resolved.LoadCancellationToken = cancellationToken;
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            var encryptedBytes = await ReadAllBytesCompatAsync(filePath, cancellationToken, resolved)
                .ConfigureAwait(false);
            return LoadEncryptedInput(encryptedBytes, password, resolved, filePath, cancellationToken);
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
            options.LoadCancellationToken = cancellationToken;
            cancellationToken.ThrowIfCancellationRequested();
            if (options?.Password is string password) return CompleteCancellableLoad(AssociateEncryptedWorkbook(await LoadEncryptedAsync(stream, password, options, cancellationToken).ConfigureAwait(false), null, password, options.AccessMode == DocumentAccessMode.ReadOnly ? null : stream, options), cancellationToken);
            ArgumentNullException.ThrowIfNull(stream);
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "workbook");
            OfficeDocumentLifecycle.EnsureSaveOnDisposeDestination(stream, resolved.PersistenceMode, nameof(stream));

            var bytes = await ReadAllBytesAsync(stream, cancellationToken, resolved)
                .ConfigureAwait(false);
            return CompleteCancellableLoad(LoadFromByteArray(
                bytes,
                resolved,
                filePath: null,
                originalStream: stream,
                leaveOriginalStreamOpen: true), cancellationToken);
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
            ArgumentNullException.ThrowIfNull(stream);
            ArgumentNullException.ThrowIfNull(password);
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            ExcelLoadOptions resolved = options ?? new ExcelLoadOptions();
            OfficeDocumentLifecycle.Validate(resolved.AccessMode, resolved.PersistenceMode, "document");
            resolved = resolved.Snapshot();
            resolved.PersistenceMode = DocumentPersistenceMode.Explicit;
            resolved.LoadCancellationToken = cancellationToken;

            var encryptedBytes = await ReadAllBytesAsync(stream, cancellationToken, resolved)
                .ConfigureAwait(false);
            return LoadEncryptedInput(encryptedBytes, password, resolved, sourcePath: null, cancellationToken);
        }

        private static ExcelDocument LoadEncryptedLegacyXls(byte[] bytes, string password, bool readOnly,
            ExcelLoadOptions options, CancellationToken cancellationToken) {
            if (!options.AllowBinaryProjection) throw new NotSupportedException("This load profile accepts XML workbook packages only; encrypted XLS projection is disabled.");
            LegacyXlsImportOptions importOptions = CreateLegacyImportOptions(options);
            importOptions.Password = password;
            importOptions.CancellationToken = cancellationToken;
            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(bytes, importOptions);
            cancellationToken.ThrowIfCancellationRequested();
            if (!workbook.WasEncryptedSource) {
                throw new InvalidDataException("LoadEncrypted requires a password-encrypted legacy XLS workbook.");
            }
            return ProjectLoadedLegacyXlsWorkbook(workbook, sourcePath: null, readOnly, cancellationToken);
        }

    }
}
