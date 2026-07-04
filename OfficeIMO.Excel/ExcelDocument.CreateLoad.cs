using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using OfficeIMO.Excel.LegacyXls;
using OfficeIMO.Excel.LegacyXls.Diagnostics;
using OfficeIMO.Excel.LegacyXls.Model;
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
        /// Creates a new Excel document at the specified path.
        /// </summary>
        /// <param name="filePath">Path to the new file.</param>
        /// <returns>Created <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Create(string filePath) {
            return Create(filePath, autoSave: true);
        }

        /// <summary>
        /// Creates a new Excel document at the specified path with explicit AutoSave behavior.
        /// </summary>
        /// <param name="filePath">Path to the new file.</param>
        /// <param name="autoSave">When true, saves changes on dispose.</param>
        /// <returns>Created <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Create(string filePath, bool autoSave) {
            // Create a spreadsheet document by supplying the filepath.
            // AutoSave controls whether workbook changes are persisted on dispose.
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook, autoSave);
            return CreateNewDocument(spreadSheetDocument, filePath, packageStream: null, sourceStream: null, copyPackageToSourceOnDispose: false, leaveSourceStreamOpen: true);
        }

        /// <summary>
        /// Creates a new Excel document in memory and optionally persists it to the provided stream on dispose.
        /// </summary>
        /// <param name="stream">Destination stream to receive the workbook package.</param>
        /// <param name="autoSave">When true, the package is written back to the stream on dispose.</param>
        /// <returns>Created <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Create(Stream stream, bool autoSave = true) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanWrite) throw new ArgumentException("Stream must be writable.", nameof(stream));

            if (autoSave && !stream.CanSeek) {
                throw new ArgumentException("Stream must support seeking when autoSave is enabled.", nameof(stream));
            }

            Stream packageStream = autoSave
                ? new NonDisposingMemoryStream(StreamBufferSize)
                : new MemoryStream(StreamBufferSize);

            var spreadSheetDocument = SpreadsheetDocument.Create(packageStream, SpreadsheetDocumentType.Workbook, false);
            return CreateNewDocument(spreadSheetDocument, filePath: null, packageStream, stream, autoSave, leaveSourceStreamOpen: true);
        }

        private static ExcelDocument CreateNewDocument(
            SpreadsheetDocument spreadSheetDocument,
            string? filePath,
            Stream? packageStream,
            Stream? sourceStream,
            bool copyPackageToSourceOnDispose,
            bool leaveSourceStreamOpen,
            bool copyPackageToFilePathOnDispose = false,
            Stream? ownedOpenStream = null) {
            bool keepPackageStream = copyPackageToSourceOnDispose || copyPackageToFilePathOnDispose;
            var document = new ExcelDocument {
                FilePath = filePath ?? string.Empty,
                _spreadSheetDocument = spreadSheetDocument
            };

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadSheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            document._workBookPart = workbookpart;

            document._packageStream = keepPackageStream ? packageStream : null;
            document._sourceStream = copyPackageToSourceOnDispose ? sourceStream : null;
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
            byte[]? unchangedPackageBytes = null) {
            bool keepPackageStream = copyPackageToSourceOnDispose || copyPackageToFilePathOnDispose;
            var document = new ExcelDocument {
                FilePath = filePath ?? string.Empty,
                _spreadSheetDocument = spreadSheetDocument,
                _workBookPart = GetWorkbookPartOrThrow(spreadSheetDocument),
                _packageStream = keepPackageStream ? packageStream : null,
                _sourceStream = copyPackageToSourceOnDispose ? sourceStream : null,
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
            bool readOnly,
            bool autoSave,
            string? filePath,
            Action<string, Exception>? log,
            OpenSettings? openSettings,
            bool preferFilePathOnFallback,
            Stream? originalStream = null,
            bool copyBackToSource = false,
            bool leaveOriginalStreamOpen = true) {
            if (bytes == null) throw new ArgumentNullException(nameof(bytes));

            if (ExcelDocumentLoadRouting.IsLegacyXls(bytes, filePath)) {
                return LoadLegacyXlsFromNormalFlow(bytes, readOnly, autoSave, filePath, openSettings);
            }

            var effectiveOpenSettings = CreateOpenSettings(openSettings, autoSave);
            bool shouldCopyBack = copyBackToSource && originalStream != null;
            bool shouldCopyBackToFilePath = !shouldCopyBack && !string.IsNullOrEmpty(filePath) && ShouldCopyBackToSource(readOnly, autoSave, openSettings);
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
                return CreateDocument(
                    memDoc,
                    filePath,
                    shouldRetainPackageStream ? normalizedStream : null,
                    shouldCopyBack ? originalStream : null,
                    shouldCopyBack,
                    leaveOriginalStreamOpen,
                    copyPackageToFilePathOnDispose: shouldCopyBackToFilePath,
                    packageContentTypesKnownNormalized: true,
                    unchangedPackageBytes: unchangedPackageBytes);
            } catch (Exception ex) when (ex is InvalidDataException || ex is OpenXmlPackageException || ex is XmlException) {
                normalizedStream?.Dispose();
                var contextMessage = filePath != null
                    ? $"Failed to open '{filePath}' after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'."
                    : "Failed to open workbook stream after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.";
                log?.Invoke($"{contextMessage} Inner exception: {ex.Message}", ex);
                throw new IOException($"{contextMessage} See inner exception for details.", ex);
            } catch {
                DisposeStream(normalizedStream);
            }

            if (preferFilePathOnFallback && !string.IsNullOrEmpty(filePath)) {
                var safePath = filePath!; // guarded by IsNullOrEmpty above
                var spreadSheetDocument = SpreadsheetDocument.Open(safePath, !readOnly, effectiveOpenSettings);
                return CreateDocument(spreadSheetDocument, filePath);
            } else {
                var fallbackStream = shouldRetainPackageStream
                    ? new NonDisposingMemoryStream(bytes.Length + StreamBufferSize)
                    : new MemoryStream(bytes.Length + StreamBufferSize);
                fallbackStream.Write(bytes, 0, bytes.Length);
                fallbackStream.Position = 0;
                var spreadSheetDocument = SpreadsheetDocument.Open(fallbackStream, !readOnly, effectiveOpenSettings);
                return CreateDocument(
                    spreadSheetDocument,
                    filePath,
                    shouldRetainPackageStream ? fallbackStream : null,
                    shouldCopyBack ? originalStream : null,
                    shouldCopyBack,
                    leaveOriginalStreamOpen,
                    copyPackageToFilePathOnDispose: shouldCopyBackToFilePath,
                    packageContentTypesKnownNormalized: false);
            }
        }

        private static byte[] ReadAllBytes(Stream stream) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer, StreamCopyBufferSize);
            return buffer.ToArray();
        }

        private static async Task<byte[]> ReadAllBytesAsync(Stream stream, CancellationToken cancellationToken) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            if (stream.CanSeek) {
                stream.Seek(0, SeekOrigin.Begin);
            }

            using var buffer = new MemoryStream();
            await stream.CopyToAsync(buffer, StreamCopyBufferSize, cancellationToken).ConfigureAwait(false);
            return buffer.ToArray();
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

        private static bool ShouldCopyBackToSource(bool readOnly, bool autoSave, OpenSettings? openSettings) {
            if (readOnly) {
                return false;
            }

            if (autoSave) {
                return true;
            }

            return openSettings?.AutoSave == true;
        }

        private static ExcelDocument LoadLegacyXlsFromNormalFlow(
            byte[] bytes,
            bool readOnly,
            bool autoSave,
            string? filePath,
            OpenSettings? openSettings) {
            if (ShouldCopyBackToSource(readOnly, autoSave, openSettings)) {
                throw new NotSupportedException("Auto-save is not supported when loading legacy binary .xls files. Save to a new .xlsx path instead.");
            }

            LegacyXlsWorkbook workbook = LegacyXlsWorkbook.Load(bytes, new LegacyXlsImportOptions {
                ReportUnsupportedRecords = true
            });
            LegacyXlsImportDiagnostic[] errors = workbook.Diagnostics
                .Where(diagnostic => diagnostic.Severity == LegacyXlsDiagnosticSeverity.Error)
                .ToArray();
            if (errors.Length > 0) {
                throw new InvalidDataException("Legacy XLS import failed: " + FormatLegacyXlsDiagnostics(errors));
            }

            return ProjectLoadedLegacyXlsWorkbook(workbook, filePath);
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

        /// <summary>
        /// Loads an existing Excel document.
        /// </summary>
        /// <param name="filePath">Path to the file.</param>
        /// <param name="readOnly">Open the file in read-only mode.</param>
        /// <param name="autoSave">Enable auto-save on dispose.</param>
        /// <param name="log">Optional callback invoked when normalization failures are encountered.</param>
        /// <param name="openSettings">Optional Open XML settings to control how the package is opened.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Load(string filePath, bool readOnly = false, bool autoSave = false, Action<string, Exception>? log = null, OpenSettings? openSettings = null) {
            if (filePath == null) {
                throw new ArgumentNullException(nameof(filePath));
            }

            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            var bytes = ReadAllBytesCompatAsync(filePath, CancellationToken.None).GetAwaiter().GetResult();
            return LoadFromByteArray(bytes, readOnly, autoSave, filePath, log, openSettings, preferFilePathOnFallback: true);
        }

        /// <summary>
        /// Loads a password-encrypted Office Open XML workbook or legacy binary `.xls` workbook.
        /// </summary>
        /// <param name="filePath">Path to the encrypted workbook.</param>
        /// <param name="password">Password used to decrypt the workbook package.</param>
        /// <param name="readOnly">Open the decrypted workbook in read-only mode.</param>
        /// <param name="autoSave">Encrypted loads do not support auto-save. Use <see cref="SaveEncrypted(string,string,bool,ExcelSaveOptions?)"/> to persist encrypted changes.</param>
        /// <param name="log">Optional callback invoked when normalization failures are encountered.</param>
        /// <param name="openSettings">Optional Open XML settings to control how the package is opened.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument LoadEncrypted(string filePath, string password, bool readOnly = false, bool autoSave = false, Action<string, Exception>? log = null, OpenSettings? openSettings = null) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            EnsureEncryptedLoadDoesNotAutoSave(autoSave, openSettings);
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            if (ExcelDocumentLoadRouting.HasLegacyXlsExtension(filePath)) {
                return LoadLegacyXls(filePath, new LegacyXlsImportOptions {
                    Password = password,
                    ReportUnsupportedRecords = true
                });
            }

            var encryptedBytes = ReadAllBytesCompatAsync(filePath, CancellationToken.None).GetAwaiter().GetResult();
            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return LoadFromByteArray(packageBytes, readOnly, autoSave: false, filePath: null, log, openSettings, preferFilePathOnFallback: false);
        }

        /// <summary>
        /// Loads an existing Excel document from the provided stream.
        /// </summary>
        /// <param name="stream">Input stream containing the workbook package.</param>
        /// <param name="readOnly">Open the document in read-only mode.</param>
        /// <param name="autoSave">Enable auto-save on dispose.</param>
        /// <param name="openSettings">Optional Open XML settings to control how the package is opened.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Load(Stream stream, bool readOnly = false, bool autoSave = false, OpenSettings? openSettings = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            bool shouldCopyBack = ShouldCopyBackToSource(readOnly, autoSave, openSettings);
            if (shouldCopyBack) {
                if (!stream.CanWrite) {
                    throw new ArgumentException("Stream must be writable when autoSave is enabled for editable documents.", nameof(stream));
                }
                if (!stream.CanSeek) {
                    throw new ArgumentException("Stream must support seeking when autoSave is enabled for editable documents.", nameof(stream));
                }
            }

            var bytes = ReadAllBytes(stream);
            return LoadFromByteArray(
                bytes,
                readOnly,
                autoSave,
                filePath: null,
                log: null,
                openSettings,
                preferFilePathOnFallback: false,
                originalStream: shouldCopyBack ? stream : null,
                copyBackToSource: shouldCopyBack,
                leaveOriginalStreamOpen: true);
        }

        /// <summary>
        /// Loads a password-encrypted Office Open XML workbook from a stream.
        /// </summary>
        /// <param name="stream">Input stream containing the encrypted workbook.</param>
        /// <param name="password">Password used to decrypt the workbook package.</param>
        /// <param name="readOnly">Open the decrypted workbook in read-only mode.</param>
        /// <param name="autoSave">Encrypted loads do not support auto-save. Use <see cref="SaveEncrypted(Stream,string,ExcelSaveOptions?)"/> to persist encrypted changes.</param>
        /// <param name="openSettings">Optional Open XML settings to control how the package is opened.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument LoadEncrypted(Stream stream, string password, bool readOnly = false, bool autoSave = false, OpenSettings? openSettings = null) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            EnsureEncryptedLoadDoesNotAutoSave(autoSave, openSettings);

            var encryptedBytes = ReadAllBytes(stream);
            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return LoadFromByteArray(packageBytes, readOnly, autoSave: false, filePath: null, log: null, openSettings, preferFilePathOnFallback: false);
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
        /// <param name="readOnly">Open the file in read-only mode.</param>
        /// <param name="autoSave">Enable auto-save on dispose.</param>
        /// <param name="openSettings">Optional Open XML settings to control how the package is opened.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        /// <exception cref="FileNotFoundException">Thrown when the file does not exist.</exception>
        public static async Task<ExcelDocument> LoadAsync(string filePath, bool readOnly = false, bool autoSave = false, OpenSettings? openSettings = null) {
            if (filePath == null) {
                throw new ArgumentNullException(nameof(filePath));
            }
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            var bytes = await ReadAllBytesCompatAsync(filePath, CancellationToken.None).ConfigureAwait(false);
            return LoadFromByteArray(bytes, readOnly, autoSave, filePath, log: null, openSettings, preferFilePathOnFallback: true);
        }

        /// <summary>
        /// Asynchronously loads a password-encrypted Office Open XML workbook.
        /// </summary>
        /// <param name="filePath">Path to the encrypted workbook.</param>
        /// <param name="password">Password used to decrypt the workbook package.</param>
        /// <param name="readOnly">Open the decrypted workbook in read-only mode.</param>
        /// <param name="autoSave">Encrypted loads do not support auto-save. Use <see cref="SaveEncrypted(string,string,bool,ExcelSaveOptions?)"/> to persist encrypted changes.</param>
        /// <param name="openSettings">Optional Open XML settings to control how the package is opened.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static async Task<ExcelDocument> LoadEncryptedAsync(string filePath, string password, bool readOnly = false, bool autoSave = false, OpenSettings? openSettings = null) {
            if (filePath == null) throw new ArgumentNullException(nameof(filePath));
            if (password == null) throw new ArgumentNullException(nameof(password));
            EnsureEncryptedLoadDoesNotAutoSave(autoSave, openSettings);
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            var encryptedBytes = await ReadAllBytesCompatAsync(filePath, CancellationToken.None).ConfigureAwait(false);
            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return LoadFromByteArray(packageBytes, readOnly, autoSave: false, filePath: null, log: null, openSettings, preferFilePathOnFallback: false);
        }

        /// <summary>
        /// Asynchronously loads an Excel document from the provided stream.
        /// </summary>
        /// <param name="stream">Input stream containing the workbook package.</param>
        /// <param name="readOnly">Open the document in read-only mode.</param>
        /// <param name="autoSave">Enable auto-save on dispose.</param>
        /// <param name="openSettings">Optional Open XML settings to control how the package is opened.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static async Task<ExcelDocument> LoadAsync(Stream stream, bool readOnly = false, bool autoSave = false, OpenSettings? openSettings = null, CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));

            bool shouldCopyBack = ShouldCopyBackToSource(readOnly, autoSave, openSettings);
            if (shouldCopyBack) {
                if (!stream.CanWrite) {
                    throw new ArgumentException("Stream must be writable when autoSave is enabled for editable documents.", nameof(stream));
                }
                if (!stream.CanSeek) {
                    throw new ArgumentException("Stream must support seeking when autoSave is enabled for editable documents.", nameof(stream));
                }
            }

            var bytes = await ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false);
            return LoadFromByteArray(
                bytes,
                readOnly,
                autoSave,
                filePath: null,
                log: null,
                openSettings,
                preferFilePathOnFallback: false,
                originalStream: shouldCopyBack ? stream : null,
                copyBackToSource: shouldCopyBack,
                leaveOriginalStreamOpen: true);
        }

        /// <summary>
        /// Asynchronously loads a password-encrypted Office Open XML workbook from a stream.
        /// </summary>
        /// <param name="stream">Input stream containing the encrypted workbook.</param>
        /// <param name="password">Password used to decrypt the workbook package.</param>
        /// <param name="readOnly">Open the decrypted workbook in read-only mode.</param>
        /// <param name="autoSave">Encrypted loads do not support auto-save. Use <see cref="SaveEncrypted(Stream,string,ExcelSaveOptions?)"/> to persist encrypted changes.</param>
        /// <param name="openSettings">Optional Open XML settings to control how the package is opened.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static async Task<ExcelDocument> LoadEncryptedAsync(Stream stream, string password, bool readOnly = false, bool autoSave = false, OpenSettings? openSettings = null, CancellationToken cancellationToken = default) {
            if (stream == null) throw new ArgumentNullException(nameof(stream));
            if (password == null) throw new ArgumentNullException(nameof(password));
            if (!stream.CanRead) throw new ArgumentException("Stream must be readable.", nameof(stream));
            EnsureEncryptedLoadDoesNotAutoSave(autoSave, openSettings);

            var encryptedBytes = await ReadAllBytesAsync(stream, cancellationToken).ConfigureAwait(false);
            var packageBytes = OfficeEncryption.DecryptPackage(encryptedBytes, password);
            return LoadFromByteArray(packageBytes, readOnly, autoSave: false, filePath: null, log: null, openSettings, preferFilePathOnFallback: false);
        }

        private static void EnsureEncryptedLoadDoesNotAutoSave(bool autoSave, OpenSettings? openSettings) {
            if (autoSave || openSettings?.AutoSave == true) {
                throw new NotSupportedException("Auto-save is not supported for encrypted Excel loads. Use SaveEncrypted to persist encrypted changes.");
            }
        }
    }
}
