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
using System.IO;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents an Excel document and provides methods for creating,
    /// loading and saving spreadsheets.
    /// </summary>
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {
        private const int StreamBufferSize = 4096;
        private static readonly System.Text.RegularExpressions.Regex _multipleUnderscoresRegex =
            new System.Text.RegularExpressions.Regex("_+", System.Text.RegularExpressions.RegexOptions.Compiled);

        private static readonly Lazy<byte[]> DefaultThemeBytes = new(() => LoadEmbeddedResource("OfficeIMO.Excel.Resources.theme1.xml"));
        // Allocated only when an operation actually needs a serialized apply stage
        internal ReaderWriterLockSlim? _lock;
        internal List<UInt32Value> id = new List<UInt32Value>() { 0 };
        private readonly Dictionary<string, int> _sharedStringCache = new Dictionary<string, int>();
        private readonly object _sharedStringLock = new object();
        // Workbook-level cache of table names for fast uniqueness checks
        private HashSet<string>? _tableNameCache;
        private System.Collections.Generic.IEqualityComparer<string> _tableNameComparer = System.StringComparer.OrdinalIgnoreCase;
        private List<ExcelSheet>? _cachedSheets;
        private bool _sheetCacheDirty = true;

        /// <summary>
        /// Enables caching of <see cref="ExcelSheet"/> wrappers for faster repeat access at the cost of higher memory usage.
        /// Set to <see langword="false"/> to avoid holding references to every sheet in very large workbooks.
        /// </summary>
        public bool SheetCachingEnabled { get; set; } = true;

        /// <summary>
        /// Controls how workbook-level table name uniqueness is compared.
        /// Defaults to <see cref="StringComparer.OrdinalIgnoreCase"/>. Changing this will reset the
        /// internal cache and rebuild it on next use. Set it once before adding tables for predictable behavior.
        /// </summary>
        public System.Collections.Generic.IEqualityComparer<string> TableNameComparer {
            get => _tableNameComparer;
            set {
                if (value == null) throw new System.ArgumentNullException(nameof(value));
                if (!object.ReferenceEquals(_tableNameComparer, value)) {
                    _tableNameComparer = value;
                    _tableNameCache = null; // rebuild lazily on next use with the new comparer
                }
            }
        }

        /// <summary>
        /// Optional default chart style preset applied to charts created in this workbook.
        /// </summary>
        public ExcelChartStylePreset? DefaultChartStylePreset { get; set; }

        /// <summary>
        /// Execution policy for controlling parallel vs sequential operations.
        /// </summary>
        public ExecutionPolicy Execution { get; } = new();

        // Default strategy mirrors CoerceValueHelper's behaviour and uses LocalDateTime so that
        // serial values are aligned with Excel's local time interpretation.
        private Func<DateTimeOffset, DateTime> _dateTimeOffsetWriteStrategy = static dto => dto.LocalDateTime;

        /// <summary>
        /// Controls how <see cref="DateTimeOffset"/> values are converted to <see cref="DateTime"/>
        /// before being written to worksheet cells. Defaults to <see cref="DateTimeOffset.LocalDateTime"/>.
        /// </summary>
        /// <remarks>
        /// The delegate influences the numeric serial value stored in the cell but does not automatically
        /// change number formats. Apply the desired cell formatting separately.
        /// </remarks>
        public Func<DateTimeOffset, DateTime> DateTimeOffsetWriteStrategy {
            get => _dateTimeOffsetWriteStrategy;
            set => _dateTimeOffsetWriteStrategy = value ?? throw new ArgumentNullException(nameof(value));
        }

        internal ReaderWriterLockSlim EnsureLock()
            => _lock ??= new ReaderWriterLockSlim(); // default: NoRecursion

        internal void EnsureWorkbookThemeAndStyles() {
            var workbookPart = _spreadSheetDocument?.WorkbookPart ?? _workBookPart;

            if (!workbookPart.GetPartsOfType<ThemePart>().Any()) {
                ThemePart themePart = workbookPart.AddNewPart<ThemePart>();
                using var themeStream = new MemoryStream(DefaultThemeBytes.Value);
                themePart.FeedData(themeStream);
            }

            var stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            if (stylesPart.Stylesheet == null) {
                stylesPart.Stylesheet = CreateDefaultStylesheet();
                stylesPart.Stylesheet.Save();
            }
        }

        private static Stylesheet CreateDefaultStylesheet() {
            var stylesheet = new Stylesheet();

            stylesheet.Fonts = new Fonts(new Font());
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

            stylesheet.Fills = new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 })
            );
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

            stylesheet.Borders = new Borders(new Border());
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

            stylesheet.CellStyleFormats = new CellStyleFormats(new CellFormat());
            stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

            stylesheet.CellFormats = new CellFormats(new CellFormat());
            stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();

            return stylesheet;
        }

        private static byte[] LoadEmbeddedResource(string resourceName) {
            var assembly = typeof(ExcelDocument).Assembly;
            using Stream? stream = assembly.GetManifestResourceStream(resourceName);
            if (stream == null) {
                throw new InvalidOperationException($"Missing embedded resource '{resourceName}'.");
            }

            using var buffer = new MemoryStream();
            stream.CopyTo(buffer);
            return buffer.ToArray();
        }

        private void MarkSheetCacheDirty()
        {
            _sheetCacheDirty = true;
            _cachedSheets = null;
        }

        private List<Sheet> ReadSheetElements()
        {
            var sheets = _spreadSheetDocument?.WorkbookPart?.Workbook.Sheets;
            if (sheets == null) {
                return new List<Sheet>();
            }

            return sheets.OfType<Sheet>().ToList();
        }

        private void UpdateSheetIdCache(List<Sheet> elements)
        {
            id.Clear();
            id.Add(0);
            foreach (Sheet s in elements) {
                var sheetId = s.SheetId;
                if (sheetId != null && !id.Contains(sheetId)) {
                    id.Add(sheetId);
                }
            }
        }

        private List<ExcelSheet> MaterializeSheets(List<Sheet> elements)
        {
            List<ExcelSheet> listExcel = new List<ExcelSheet>(elements.Count);
            foreach (Sheet s in elements) {
                listExcel.Add(new ExcelSheet(this, _spreadSheetDocument!, s));
            }

            return listExcel;
        }

        private void RebuildSheetCacheLocked()
        {
            var elements = ReadSheetElements();
            UpdateSheetIdCache(elements);
            _cachedSheets = SheetCachingEnabled ? MaterializeSheets(elements) : null;
            _sheetCacheDirty = false;
        }

        private void EnsureSheetCacheInitialized(ReaderWriterLockSlim? lck)
        {
            if (!(_sheetCacheDirty || _cachedSheets == null)) return;

            if (Locking.IsNoLock || lck is null || lck.IsWriteLockHeld) {
                RebuildSheetCacheLocked();
                return;
            }

            lck.EnterWriteLock();
            try {
                if (_sheetCacheDirty || _cachedSheets == null) {
                    RebuildSheetCacheLocked();
                }
            } finally {
                lck.ExitWriteLock();
            }
        }

        private List<ExcelSheet> CloneSheetCache()
        {
            if (_cachedSheets == null) {
                return new List<ExcelSheet>();
            }

            return new List<ExcelSheet>(_cachedSheets);
        }

        private List<ExcelSheet> BuildSheetsWithoutCaching()
        {
            var elements = ReadSheetElements();
            UpdateSheetIdCache(elements);
            return MaterializeSheets(elements);
        }

        internal void InvalidateSheetCache()
        {
            Locking.ExecuteWrite(EnsureLock(), MarkSheetCacheDirty);
        }

        /// <summary>
        /// Gets a list of worksheets contained in the document.
        /// </summary>
        public List<ExcelSheet> Sheets {
            get {
                var lck = EnsureLock();
                if (Locking.IsNoLock || lck is null) {
                    if (SheetCachingEnabled) {
                        EnsureSheetCacheInitialized(lck);
                        return CloneSheetCache();
                    }

                    return BuildSheetsWithoutCaching();
                }

                if (!SheetCachingEnabled) {
                    lck.EnterReadLock();
                    try {
                        return BuildSheetsWithoutCaching();
                    } finally {
                        lck.ExitReadLock();
                    }
                }

                lck.EnterReadLock();
                try {
                    if (!(_sheetCacheDirty || _cachedSheets == null)) {
                        return CloneSheetCache();
                    }
                } finally {
                    lck.ExitReadLock();
                }

                lck.EnterUpgradeableReadLock();
                try {
                    if (_sheetCacheDirty || _cachedSheets == null) {
                        EnsureSheetCacheInitialized(lck);
                    }

                    return CloneSheetCache();
                } finally {
                    lck.ExitUpgradeableReadLock();
                }
            }
        }

        /// <summary>
        /// Underlying Open XML spreadsheet document instance.
        /// </summary>
        public SpreadsheetDocument _spreadSheetDocument = null!;
        private WorkbookPart _workBookPart = null!;
        private SharedStringTablePart? _sharedStringTablePart;
        private Stream? _packageStream;
        private Stream? _sourceStream;
        private bool _copyPackageToSourceOnDispose;
        private bool _leaveSourceStreamOpen = true;

        private const int StreamCopyBufferSize = 81920;

        private static async Task<byte[]> ReadAllBytesCompatAsync(string path, CancellationToken ct) {
#if NETSTANDARD2_0 || NET472 || NET48
            using (var fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read, 8192, FileOptions.Asynchronous))
            {
                var mem = new MemoryStream((int)Math.Max(0, fs.Length) + 8192);
                await fs.CopyToAsync(mem, 81920, ct).ConfigureAwait(false);
                return mem.ToArray();
            }
#else
            return await File.ReadAllBytesAsync(path, ct).ConfigureAwait(false);
#endif
        }

        private static OpenSettings CreateOpenSettings(OpenSettings? openSettings, bool autoSave) {
            bool shouldAutoSave = autoSave || (openSettings?.AutoSave ?? false);

            if (openSettings is null) {
                return new OpenSettings { AutoSave = shouldAutoSave };
            }

            if (openSettings.AutoSave == shouldAutoSave) {
                return openSettings;
            }

            return new OpenSettings {
                AutoSave = shouldAutoSave,
                CompatibilityLevel = openSettings.CompatibilityLevel,
                MarkupCompatibilityProcessSettings = openSettings.MarkupCompatibilityProcessSettings,
                MaxCharactersInPart = openSettings.MaxCharactersInPart,
            };
        }

        /// <summary>
        /// Path to the file backing this document.
        /// </summary>
        public string FilePath = string.Empty;

        /// <summary>
        /// Built-in (core) document properties (Title, Creator, etc.).
        /// </summary>
        public BuiltinDocumentProperties BuiltinDocumentProperties = null!;

        /// <summary>
        /// Extended (application) properties (Company, Manager, etc.).
        /// </summary>
        public ApplicationProperties ApplicationProperties = null!;

        /// <summary>
        /// FileOpenAccess of the document
        /// </summary>
        public FileAccess FileOpenAccess => _spreadSheetDocument.FileOpenAccess;

        /// <summary>
        /// Indicates whether the document is valid.
        /// </summary>
        public bool DocumentIsValid {
            get {
                if (DocumentValidationErrors.Count > 0) {
                    return false;
                }

                return true;
            }
        }

        /// <summary>
        /// Gets the list of validation errors for the document.
        /// </summary>
        public List<ValidationErrorInfo> DocumentValidationErrors {
            get {
                return ValidateDocument();
            }
        }

        /// <summary>
        /// Returns the workbook-level cache of table names, initializing it from the current
        /// document if needed. Case-insensitive comparison.
        /// </summary>
        internal HashSet<string> GetOrInitTableNameCache() {
            // Fast path without locking
            if (_tableNameCache != null) return _tableNameCache;

            // Initialize without taking a new lock if we're already in a write scope
            if (Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld)) {
                if (_tableNameCache == null) {
                    var set = new HashSet<string>(_tableNameComparer);
                    var wb = _spreadSheetDocument.WorkbookPart;
                    if (wb != null) {
                        foreach (var ws in wb.WorksheetParts) {
                            foreach (var tdp in ws.TableDefinitionParts) {
                                var n = tdp.Table?.Name?.Value;
                                if (!string.IsNullOrEmpty(n)) set.Add(n!);
                            }
                        }
                    }
                    _tableNameCache = set;
                }
                return _tableNameCache!;
            }

            // Otherwise, use write lock for thread safety
            return Locking.ExecuteWrite(EnsureLock(), () => {
                if (_tableNameCache != null) return _tableNameCache;
                var set = new HashSet<string>(_tableNameComparer);
                var wb = _spreadSheetDocument.WorkbookPart;
                if (wb != null) {
                    foreach (var ws in wb.WorksheetParts) {
                        foreach (var tdp in ws.TableDefinitionParts) {
                            var n = tdp.Table?.Name?.Value;
                            if (!string.IsNullOrEmpty(n)) set.Add(n!);
                        }
                    }
                }
                _tableNameCache = set;
                return _tableNameCache;
            });
        }

        /// <summary>
        /// Adds the given table name to the cache. Should be called once the name is finalized.
        /// </summary>
        internal void ReserveTableName(string name) {
            if (string.IsNullOrWhiteSpace(name)) return;
            var cache = GetOrInitTableNameCache();
            cache.Add(name);
        }

        /// <summary>
        /// Removes the given table name from the cache. Intended for future table deletion APIs.
        /// Safe to call even if the cache hasn't been initialized.
        /// </summary>
        internal void RemoveReservedTableName(string name) {
            if (string.IsNullOrWhiteSpace(name)) return;
            if (_tableNameCache == null) return;
            _tableNameCache.Remove(name);
        }

        /// <summary>
        /// Validates the document using the specified file format version.
        /// </summary>
        /// <param name="fileFormatVersions">File format version to validate against.</param>
        /// <returns>List of validation errors.</returns>
        public List<ValidationErrorInfo> ValidateDocument(FileFormatVersions fileFormatVersions = FileFormatVersions.Microsoft365) {
            List<ValidationErrorInfo> listErrors = new List<ValidationErrorInfo>();
            OpenXmlValidator validator = new OpenXmlValidator(fileFormatVersions);
            foreach (ValidationErrorInfo error in validator.Validate(_spreadSheetDocument)) {
                listErrors.Add(error);
            }
            return listErrors;
        }

        internal SharedStringTablePart SharedStringTablePart {
            get {
                // Check if already initialized without lock first (double-check locking pattern)
                if (_sharedStringTablePart != null) {
                    return _sharedStringTablePart;
                }

                // Check if we're in a NoLock scope or already have a lock - if so, initialize without locking
                if (Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld)) {
                    if (_sharedStringTablePart == null) {
                        if (_workBookPart.GetPartsOfType<SharedStringTablePart>().Any()) {
                            _sharedStringTablePart = _workBookPart.GetPartsOfType<SharedStringTablePart>().First();
                        } else {
                            _sharedStringTablePart = _workBookPart.AddNewPart<SharedStringTablePart>();
                            _sharedStringTablePart.SharedStringTable = new SharedStringTable();
                        }
                    }
                    return _sharedStringTablePart;
                }

                // Use write lock for initialization when no lock is held
                return Locking.ExecuteWrite(EnsureLock(), () => {
                    // Double-check inside the lock
                    if (_sharedStringTablePart == null) {
                        if (_workBookPart.GetPartsOfType<SharedStringTablePart>().Any()) {
                            _sharedStringTablePart = _workBookPart.GetPartsOfType<SharedStringTablePart>().First();
                        } else {
                            _sharedStringTablePart = _workBookPart.AddNewPart<SharedStringTablePart>();
                            _sharedStringTablePart.SharedStringTable = new SharedStringTable();
                        }
                    }
                    return _sharedStringTablePart;
                });
            }
        }

        internal int GetSharedStringIndex(string text) {
            lock (_sharedStringLock) {
                // Check cache first
                if (_sharedStringCache.TryGetValue(text, out int cachedIndex)) {
                    return cachedIndex;
                }

                var sharedStringTable = SharedStringTablePart.SharedStringTable;

                // If cache is empty, rebuild it
                if (_sharedStringCache.Count == 0) {
                    int idx = 0;
                    foreach (SharedStringItem item in sharedStringTable.Elements<SharedStringItem>()) {
                        _sharedStringCache[item.InnerText] = idx;
                        idx++;
                    }

                    // Check again after rebuilding cache
                    if (_sharedStringCache.TryGetValue(text, out int foundIndex)) {
                        return foundIndex;
                    }
                }

                // Add new string
                int newIndex = _sharedStringCache.Count;
                sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
                sharedStringTable.Save();
                _sharedStringCache[text] = newIndex;

                return newIndex;
            }
        }

        /// <summary>
        /// Creates a new Excel document at the specified path.
        /// </summary>
        /// <param name="filePath">Path to the new file.</param>
        /// <returns>Created <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Create(string filePath) {
            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
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

            var spreadSheetDocument = SpreadsheetDocument.Create(packageStream, SpreadsheetDocumentType.Workbook, true);
            return CreateNewDocument(spreadSheetDocument, filePath: null, packageStream, stream, autoSave, leaveSourceStreamOpen: true);
        }

        private static ExcelDocument CreateNewDocument(
            SpreadsheetDocument spreadSheetDocument,
            string? filePath,
            Stream? packageStream,
            Stream? sourceStream,
            bool copyPackageToSourceOnDispose,
            bool leaveSourceStreamOpen) {
            var document = new ExcelDocument {
                FilePath = filePath ?? string.Empty,
                _spreadSheetDocument = spreadSheetDocument
            };

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadSheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();
            document._workBookPart = workbookpart;

            document._packageStream = copyPackageToSourceOnDispose ? packageStream : null;
            document._sourceStream = copyPackageToSourceOnDispose ? sourceStream : null;
            document._copyPackageToSourceOnDispose = copyPackageToSourceOnDispose && sourceStream != null;
            document._leaveSourceStreamOpen = leaveSourceStreamOpen;

            // Initialize document property helpers
            document.BuiltinDocumentProperties = new BuiltinDocumentProperties(document);
            document.ApplicationProperties = new ApplicationProperties(document);

            return document;
        }
        private static ExcelDocument CreateDocument(
            SpreadsheetDocument spreadSheetDocument,
            string? filePath,
            Stream? packageStream = null,
            Stream? sourceStream = null,
            bool copyPackageToSourceOnDispose = false,
            bool leaveSourceStreamOpen = true) {
            var document = new ExcelDocument {
                FilePath = filePath ?? string.Empty,
                _spreadSheetDocument = spreadSheetDocument,
                _workBookPart = GetWorkbookPartOrThrow(spreadSheetDocument),
                _packageStream = copyPackageToSourceOnDispose ? packageStream : null,
                _sourceStream = copyPackageToSourceOnDispose ? sourceStream : null,
                _copyPackageToSourceOnDispose = copyPackageToSourceOnDispose && sourceStream != null,
                _leaveSourceStreamOpen = leaveSourceStreamOpen,
            };

            document.BuiltinDocumentProperties = new BuiltinDocumentProperties(document);
            document.ApplicationProperties = new ApplicationProperties(document);
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

            var effectiveOpenSettings = CreateOpenSettings(openSettings, autoSave);
            bool shouldCopyBack = copyBackToSource && originalStream != null;

            MemoryStream? normalizedStream = null;

            try {
                normalizedStream = shouldCopyBack
                    ? new NonDisposingMemoryStream(bytes.Length + StreamBufferSize)
                    : new MemoryStream(bytes.Length + StreamBufferSize);
                normalizedStream.Write(bytes, 0, bytes.Length);
                normalizedStream.Position = 0;

                Utilities.ExcelPackageUtilities.NormalizeContentTypes(normalizedStream, leaveOpen: true);
                normalizedStream.Position = 0;

                var memDoc = SpreadsheetDocument.Open(normalizedStream, !readOnly, effectiveOpenSettings);
                return CreateDocument(
                    memDoc,
                    filePath,
                    shouldCopyBack ? normalizedStream : null,
                    shouldCopyBack ? originalStream : null,
                    shouldCopyBack,
                    leaveOriginalStreamOpen);
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
                var fallbackStream = shouldCopyBack
                    ? new NonDisposingMemoryStream(bytes.Length + StreamBufferSize)
                    : new MemoryStream(bytes.Length + StreamBufferSize);
                fallbackStream.Write(bytes, 0, bytes.Length);
                fallbackStream.Position = 0;
                var spreadSheetDocument = SpreadsheetDocument.Open(fallbackStream, !readOnly, effectiveOpenSettings);
                return CreateDocument(
                    spreadSheetDocument,
                    filePath,
                    shouldCopyBack ? fallbackStream : null,
                    shouldCopyBack ? originalStream : null,
                    shouldCopyBack,
                    leaveOriginalStreamOpen);
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

            var effectiveOpenSettings = CreateOpenSettings(openSettings, autoSave);

            // Try direct file streaming first for better memory efficiency and avoid large intermediate buffers.
            // Packages with content type issues may require normalization, so fall back to buffered reads on failures.
            try {
                var spreadSheetDocument = SpreadsheetDocument.Open(filePath, !readOnly, effectiveOpenSettings);
                return CreateDocument(spreadSheetDocument, filePath);
            } catch (Exception ex) when (ex is InvalidDataException || ex is OpenXmlPackageException || ex is XmlException) {
                log?.Invoke($"Failed to open '{filePath}' directly. Falling back to normalized stream. Inner exception: {ex.Message}", ex);
            }

            var bytes = ReadAllBytesCompatAsync(filePath, CancellationToken.None).GetAwaiter().GetResult();
            return LoadFromByteArray(bytes, readOnly, autoSave, filePath, log, openSettings, preferFilePathOnFallback: true);
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
        /// Creates a new Excel document with a single worksheet.
        /// </summary>
        /// <param name="filePath">Path to the new file.</param>
        /// <param name="workSheetName">Name of the worksheet.</param>
        /// <returns>Created <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Create(string filePath, string workSheetName) {
            ExcelDocument excelDocument = Create(filePath);
            // Prefer a sanitized sheet name for convenience in the common Create(path, name) flow
            excelDocument.AddWorkSheet(workSheetName, SheetNameValidationMode.Sanitize);
            return excelDocument;
        }

        /// <summary>
        /// Adds a worksheet to the document.
        /// </summary>
        /// <param name="workSheetName">Worksheet name.</param>
        /// <returns>Created <see cref="ExcelSheet"/> instance.</returns>
        public ExcelSheet AddWorkSheet(string workSheetName = "") {
            return AddWorkSheet(workSheetName, SheetNameValidationMode.None);
        }

        /// <summary>
        /// Adds a worksheet to the document with control over name validation.
        /// </summary>
        /// <param name="workSheetName">Requested worksheet name.</param>
        /// <param name="validationMode">How to validate the sheet name: None (no checks), Sanitize (coerce), or Strict (throw on invalid).</param>
        /// <returns>Created <see cref="ExcelSheet"/> instance.</returns>
        public ExcelSheet AddWorkSheet(string workSheetName, SheetNameValidationMode validationMode) {
            return Locking.ExecuteWrite(EnsureLock(), () => {
                EnsureSheetCacheInitialized(_lock);
                string name = ValidateOrSanitizeSheetName(workSheetName, validationMode);
                ExcelSheet excelSheet = new ExcelSheet(this, _workBookPart, _spreadSheetDocument, name);
                MarkSheetCacheDirty();
                return excelSheet;
            });
        }

        private string ValidateOrSanitizeSheetName(string name, SheetNameValidationMode mode) {
            // Collect existing names (case-insensitive)
            var existing = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            foreach (var s in _workBookPart.Workbook.Sheets?.OfType<DocumentFormat.OpenXml.Spreadsheet.Sheet>() ?? System.Linq.Enumerable.Empty<DocumentFormat.OpenXml.Spreadsheet.Sheet>()) {
                var existingName = s.Name?.Value;
                if (!string.IsNullOrEmpty(existingName)) existing.Add(existingName!);
            }

            if (mode == SheetNameValidationMode.None) {
                // Preserve historical behavior: default to "Sheet1" when empty
                if (string.IsNullOrEmpty(name)) name = "Sheet1";
                return name;
            }

            // Rules common to Sanitize/Strict
            static bool ContainsInvalidChars(string s) {
                foreach (char c in s) {
                    if (c == ':' || c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']') return true;
                }
                return false;
            }

            string baseName = name ?? string.Empty;
            baseName = baseName.Trim();
            baseName = baseName.Trim('\'', ' ');

            if (mode == SheetNameValidationMode.Strict) {
                if (string.IsNullOrEmpty(baseName)) throw new System.ArgumentException("Worksheet name cannot be empty.", nameof(name));
                if (baseName.Length > 31) throw new System.ArgumentException("Worksheet name cannot exceed 31 characters.", nameof(name));
                if (ContainsInvalidChars(baseName)) throw new System.ArgumentException("Worksheet name contains invalid characters (: \\ / ? * [ ]).", nameof(name));
                if (existing.Contains(baseName)) throw new System.ArgumentException($"Worksheet name '{baseName}' already exists.", nameof(name));
                return baseName;
            }

            // Sanitize
            var sb = new System.Text.StringBuilder(baseName.Length > 0 ? baseName.Length : 5);
            foreach (char c in baseName) {
                if (c == ':' || c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']') sb.Append('_');
                else sb.Append(c);
            }
            string cleaned = sb.ToString().Trim();
            // Collapse multiple underscores and trim leading/trailing underscores for nicer names
            cleaned = _multipleUnderscoresRegex.Replace(cleaned, "_");
            cleaned = cleaned.Trim('_');
            if (cleaned.Length == 0) cleaned = "Sheet";
            if (cleaned.Length > 31) cleaned = cleaned.Substring(0, 31);

            // Ensure uniqueness by appending (2), (3), ...
            string candidate = cleaned;
            int n = 2;
            while (existing.Contains(candidate)) {
                string suffix = " (" + n.ToString(System.Globalization.CultureInfo.InvariantCulture) + ")";
                int maxBase = 31 - suffix.Length;
                string basePart = cleaned.Length > maxBase ? cleaned.Substring(0, maxBase) : cleaned;
                candidate = basePart + suffix;
                n++;
            }
            return candidate;
        }

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
            foreach (var sheet in Sheets) {
                sheet.Preflight();
            }
        }

        /// <summary>
        /// Closes the underlying spreadsheet document.
        /// </summary>
        public void Close() {
            this._spreadSheetDocument.Dispose();
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
            var path = string.IsNullOrEmpty(filePath) ? FilePath : filePath;

            // Ensure target directory is writable
            if (File.Exists(path) && new FileInfo(path).IsReadOnly) {
                throw new IOException($"Failed to save to '{path}'. The file is read-only.");
            }
            EnsureDirectoryWritable(path);

            var payload = PreparePackageForSave(options);

            using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None)) {
                fs.Write(payload.PackageBytes, 0, payload.PackageBytes.Length);
                fs.Flush();
            }

            try { payload.Properties.ApplyTo(path); } catch { }
            try { ExcelPackageUtilities.NormalizeContentTypes(path); } catch { }
            FilePath = path;

            var fileBytes = File.ReadAllBytes(path);
            ReloadFromBytes(fileBytes);

            if (openExcel) {
                Helpers.Open(path, true);
            }

            if (options?.ValidateOpenXml == true) {
                var errors = ValidateOpenXml();
                if (errors.Count > 0) {
                    throw new System.InvalidOperationException("OpenXML validation failed:\n" + string.Join("\n", errors));
                }
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
            var target = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
            if (File.Exists(target) && new FileInfo(target).IsReadOnly) {
                throw new IOException($"Failed to save to '{target}'. The file is read-only.");
            }
            EnsureDirectoryWritable(target);

            var payload = PreparePackageForSave(options);

            using (var fs = new FileStream(target, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 8192, FileOptions.Asynchronous)) {
                await fs.WriteAsync(payload.PackageBytes, 0, payload.PackageBytes.Length, cancellationToken).ConfigureAwait(false);
                await fs.FlushAsync(cancellationToken).ConfigureAwait(false);
            }

            try { payload.Properties.ApplyTo(target); } catch { }
            try { ExcelPackageUtilities.NormalizeContentTypes(target); } catch { }
            FilePath = target;

            var fileBytes = await ReadAllBytesCompatAsync(target, cancellationToken).ConfigureAwait(false);
            ReloadFromBytes(fileBytes);

            if (openExcel) {
                Open(filePath, true);
            }

            if (options?.ValidateOpenXml == true) {
                var errors = ValidateOpenXml();
                if (errors.Count > 0) {
                    throw new System.InvalidOperationException("OpenXML validation failed:\n" + string.Join("\n", errors));
                }
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

            var payload = PreparePackageForSave(options);
            var withProperties = payload.Properties.ApplyTo(payload.PackageBytes);
            var finalizedBytes = NormalizePackageBytes(withProperties);
            destination.Write(finalizedBytes, 0, finalizedBytes.Length);
            try { destination.Flush(); } catch (NotSupportedException) { }

            ReloadFromBytes(finalizedBytes);

            if (options?.ValidateOpenXml == true) {
                var errors = ValidateOpenXml();
                if (errors.Count > 0) {
                    throw new System.InvalidOperationException("OpenXML validation failed:\n" + string.Join("\n", errors));
                }
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

            var payload = PreparePackageForSave(options);
            var withProperties = payload.Properties.ApplyTo(payload.PackageBytes);
            var finalizedBytes = NormalizePackageBytes(withProperties);
            await destination.WriteAsync(finalizedBytes, 0, finalizedBytes.Length, cancellationToken).ConfigureAwait(false);
            try { await destination.FlushAsync(cancellationToken).ConfigureAwait(false); } catch (NotSupportedException) { }

            ReloadFromBytes(finalizedBytes);

            if (options?.ValidateOpenXml == true) {
                var errors = ValidateOpenXml();
                if (errors.Count > 0) {
                    throw new System.InvalidOperationException("OpenXML validation failed:\n" + string.Join("\n", errors));
                }
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

        private SavePayload PreparePackageForSave(ExcelSaveOptions? options) {
            // Ensure all worksheets have up-to-date dimensions and proper element ordering before saving
            foreach (var sheet in Sheets) {
                sheet.UpdateSheetDimension();
                sheet.EnsureWorksheetElementOrder();
                sheet.Commit();
            }

            // Always preflight to remove orphaned/empty containers that can trigger Excel repairs
            try { PreflightWorkbook(); } catch { }
            if (options?.SafePreflight == true) {
                // Already performed above; branch kept for semantic clarity
            }

            if (options?.SafeRepairDefinedNames == true) {
                try { RepairDefinedNames(save: true); } catch { }
            }

            _workBookPart.Workbook.Save();
            try { _spreadSheetDocument.PackageProperties.Modified = DateTime.UtcNow; } catch { }

            PackagePropertiesSnapshot propertiesSnapshot = PackagePropertiesSnapshot.Capture(_spreadSheetDocument);

            var snapshot = new MemoryStream();
            using (_spreadSheetDocument.Clone(snapshot)) { }
            snapshot.Position = 0;

            var packageBytes = snapshot.ToArray();

            try { _spreadSheetDocument.Dispose(); } catch { }

            return new SavePayload(packageBytes, propertiesSnapshot);
        }

        private void ReloadFromBytes(byte[] packageBytes) {
            var mem = new MemoryStream(packageBytes.Length + 8192);
            mem.Write(packageBytes, 0, packageBytes.Length);
            mem.Position = 0;
            var reopenSettings = new OpenSettings { AutoSave = true };
            _spreadSheetDocument = SpreadsheetDocument.Open(mem, true, reopenSettings);
            _workBookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            _sharedStringTablePart = null;
        }

        private static byte[] NormalizePackageBytes(byte[] packageBytes) {
            var working = new MemoryStream(packageBytes.Length + StreamBufferSize);
            working.Write(packageBytes, 0, packageBytes.Length);
            working.Position = 0;

            try {
                ExcelPackageUtilities.NormalizeContentTypes(working, leaveOpen: true);
            } catch {
            }

            if (working.CanSeek) {
                working.Position = 0;
            }

            return working.ToArray();
        }

        private sealed class PackagePropertiesSnapshot {
            private readonly string? _title;
            private readonly string? _creator;
            private readonly string? _subject;
            private readonly string? _category;
            private readonly string? _description;
            private readonly string? _keywords;
            private readonly string? _lastModifiedBy;
            private readonly string? _version;
            private readonly DateTime? _created;
            private readonly DateTime? _modified;
            private readonly DateTime? _lastPrinted;

            private PackagePropertiesSnapshot(
                string? title,
                string? creator,
                string? subject,
                string? category,
                string? description,
                string? keywords,
                string? lastModifiedBy,
                string? version,
                DateTime? created,
                DateTime? modified,
                DateTime? lastPrinted) {
                _title = title;
                _creator = creator;
                _subject = subject;
                _category = category;
                _description = description;
                _keywords = keywords;
                _lastModifiedBy = lastModifiedBy;
                _version = version;
                _created = created;
                _modified = modified;
                _lastPrinted = lastPrinted;
            }

            public static PackagePropertiesSnapshot Capture(SpreadsheetDocument document) {
                try {
                    var props = document.PackageProperties;
                    return new PackagePropertiesSnapshot(
                        props.Title,
                        props.Creator,
                        props.Subject,
                        props.Category,
                        props.Description,
                        props.Keywords,
                        props.LastModifiedBy,
                        props.Version,
                        props.Created,
                        props.Modified,
                        props.LastPrinted);
                } catch {
                    return new PackagePropertiesSnapshot(null, null, null, null, null, null, null, null, null, null, null);
                }
            }

            public void ApplyTo(string packagePath) {
                if (string.IsNullOrWhiteSpace(packagePath) || !File.Exists(packagePath)) {
                    return;
                }

                try {
                    using var package = Package.Open(packagePath, FileMode.Open, FileAccess.ReadWrite);
                    var dst = package.PackageProperties;
                    dst.Title = _title;
                    dst.Creator = _creator;
                    dst.Subject = _subject;
                    dst.Category = _category;
                    dst.Description = _description;
                    dst.Keywords = _keywords;
                    dst.LastModifiedBy = _lastModifiedBy;
                    dst.Version = _version;
                    dst.Created = _created;
                    dst.Modified = _modified ?? DateTime.UtcNow;
                    dst.LastPrinted = _lastPrinted;
                } catch {
                }
            }

            public byte[] ApplyTo(byte[] packageBytes) {
                if (packageBytes == null) throw new ArgumentNullException(nameof(packageBytes));
                if (packageBytes.Length == 0) return packageBytes;

                try {
                    var working = new MemoryStream(packageBytes.Length + StreamBufferSize);
                    working.Write(packageBytes, 0, packageBytes.Length);
                    working.Position = 0;

                    using (var package = Package.Open(working, FileMode.Open, FileAccess.ReadWrite)) {
                        var dst = package.PackageProperties;
                        dst.Title = _title;
                        dst.Creator = _creator;
                        dst.Subject = _subject;
                        dst.Category = _category;
                        dst.Description = _description;
                        dst.Keywords = _keywords;
                        dst.LastModifiedBy = _lastModifiedBy;
                        dst.Version = _version;
                        dst.Created = _created;
                        dst.Modified = _modified ?? DateTime.UtcNow;
                        dst.LastPrinted = _lastPrinted;
                    }

                    if (working.CanSeek) {
                        working.Position = 0;
                    }

                    return working.ToArray();
                } catch {
                    return packageBytes;
                }
            }
        }

        private sealed class SavePayload {
            public SavePayload(byte[] packageBytes, PackagePropertiesSnapshot properties) {
                PackageBytes = packageBytes;
                Properties = properties;
            }

            public byte[] PackageBytes { get; }
            public PackagePropertiesSnapshot Properties { get; }
        }

        private bool _disposed;

        /// <summary>
        /// Releases resources used by the document.
        /// </summary>
        public void Dispose() {
            DisposeAsync().GetAwaiter().GetResult();
        }

        /// <summary>
        /// Asynchronously releases resources used by the document.
        /// </summary>
        public async ValueTask DisposeAsync() {
            if (_disposed) {
                return;
            }

            if (this._spreadSheetDocument != null) {
                try {
                    if (this._spreadSheetDocument.AutoSave && this._spreadSheetDocument.FileOpenAccess != FileAccess.Read) {
                        _workBookPart?.Workbook.Save();
                    }

                    await Task.Run(() => this._spreadSheetDocument.Dispose()).ConfigureAwait(false);
                } catch {
                    // ignored
                }
                this._spreadSheetDocument = null!;
            }

            PersistPackageToSourceIfNeeded();

            _lock?.Dispose();
            _disposed = true;
            GC.SuppressFinalize(this);
        }

        private void PersistPackageToSourceIfNeeded() {
            if (_packageStream == null) {
                return;
            }

            try {
                if (_copyPackageToSourceOnDispose && _sourceStream != null) {
                    PersistPackageToSource();
                }
            } catch {
                // ignored
            } finally {
                DisposeStream(_packageStream);

                if (_copyPackageToSourceOnDispose && _sourceStream != null) {
                    if (!_leaveSourceStreamOpen) {
                        try {
                            _sourceStream.Dispose();
                        } catch {
                            // ignored
                        }
                    } else if (_sourceStream.CanSeek) {
                        try {
                            _sourceStream.Seek(0, SeekOrigin.Begin);
                        } catch {
                            // ignored
                        }
                    }
                }

                _packageStream = null;
                _sourceStream = null;
                _copyPackageToSourceOnDispose = false;
                _leaveSourceStreamOpen = true;
            }
        }

        private void PersistPackageToSource() {
            var packageStream = _packageStream ?? throw new InvalidOperationException("Package stream is not available.");
            var targetStream = _sourceStream ?? throw new InvalidOperationException("Source stream is not available.");

            if (!targetStream.CanSeek) {
                throw new InvalidOperationException("The provided stream must support seeking when autoSave is enabled.");
            }

            if (packageStream.CanSeek) {
                packageStream.Seek(0, SeekOrigin.Begin);
            }

            targetStream.Seek(0, SeekOrigin.Begin);
            targetStream.SetLength(0);
            packageStream.CopyTo(targetStream, StreamCopyBufferSize);
            targetStream.Flush();
            targetStream.Seek(0, SeekOrigin.Begin);
        }
    }
}
