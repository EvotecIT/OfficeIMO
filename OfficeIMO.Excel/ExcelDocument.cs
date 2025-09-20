using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using OfficeIMO.Excel.Utilities;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents an Excel document and provides methods for creating,
    /// loading and saving spreadsheets.
    /// </summary>
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {
        private static readonly System.Text.RegularExpressions.Regex _multipleUnderscoresRegex =
            new System.Text.RegularExpressions.Regex("_+", System.Text.RegularExpressions.RegexOptions.Compiled);
        // Allocated only when an operation actually needs a serialized apply stage
        internal ReaderWriterLockSlim? _lock;
        internal List<UInt32Value> id = new List<UInt32Value>() { 0 };
        private readonly Dictionary<string, int> _sharedStringCache = new Dictionary<string, int>();
        private readonly object _sharedStringLock = new object();
        // Workbook-level cache of table names for fast uniqueness checks
        private HashSet<string>? _tableNameCache;
        private System.Collections.Generic.IEqualityComparer<string> _tableNameComparer = System.StringComparer.OrdinalIgnoreCase;

        /// <summary>
        /// Controls how workbook-level table name uniqueness is compared.
        /// Defaults to <see cref="StringComparer.OrdinalIgnoreCase"/>. Changing this will reset the
        /// internal cache and rebuild it on next use. Set it once before adding tables for predictable behavior.
        /// </summary>
        public System.Collections.Generic.IEqualityComparer<string> TableNameComparer
        {
            get => _tableNameComparer;
            set
            {
                if (value == null) throw new System.ArgumentNullException(nameof(value));
                if (!object.ReferenceEquals(_tableNameComparer, value))
                {
                    _tableNameComparer = value;
                    _tableNameCache = null; // rebuild lazily on next use with the new comparer
                }
            }
        }

        /// <summary>
        /// Execution policy for controlling parallel vs sequential operations.
        /// </summary>
        public ExecutionPolicy Execution { get; } = new();

        internal ReaderWriterLockSlim EnsureLock()
            => _lock ??= new ReaderWriterLockSlim(); // default: NoRecursion

        /// <summary>
        /// Gets a list of worksheets contained in the document.
        /// </summary>
        public List<ExcelSheet> Sheets {
            get {
                // Since we need to both read and write, we'll use a write lock for the entire operation
                // to avoid nested lock issues
                return Locking.ExecuteWrite(EnsureLock(), () => {
                    List<ExcelSheet> listExcel = new List<ExcelSheet>();
                    List<Sheet>? elements = null;
                    var sheets = _spreadSheetDocument?.WorkbookPart?.Workbook.Sheets;
                    if (sheets != null) {
                        elements = sheets.OfType<Sheet>().ToList();
                        foreach (Sheet s in elements) {
                            listExcel.Add(new ExcelSheet(this, _spreadSheetDocument!, s));
                        }
                    }

                    // Update the id list
                    id.Clear();
                    id.Add(0);
                    if (elements != null) {
                        foreach (Sheet s in elements) {
                            var sheetId = s.SheetId;
                            if (sheetId != null && !id.Contains(sheetId)) {
                                id.Add(sheetId);
                            }
                        }
                    }

                    return listExcel;
                });
            }
        }

        /// <summary>
        /// Underlying Open XML spreadsheet document instance.
        /// </summary>
        public SpreadsheetDocument _spreadSheetDocument = null!;
        private WorkbookPart _workBookPart = null!;
        private SharedStringTablePart? _sharedStringTablePart;

        private static async Task<byte[]> ReadAllBytesCompatAsync(string path, CancellationToken ct)
        {
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
        internal HashSet<string> GetOrInitTableNameCache()
        {
            // Fast path without locking
            if (_tableNameCache != null) return _tableNameCache;

            // Initialize without taking a new lock if we're already in a write scope
            if (Locking.IsNoLock || (_lock != null && _lock.IsWriteLockHeld))
            {
                if (_tableNameCache == null)
                {
                    var set = new HashSet<string>(_tableNameComparer);
                    var wb = _spreadSheetDocument.WorkbookPart;
                    if (wb != null)
                    {
                        foreach (var ws in wb.WorksheetParts)
                        {
                            foreach (var tdp in ws.TableDefinitionParts)
                            {
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
            return Locking.ExecuteWrite(EnsureLock(), () =>
            {
                if (_tableNameCache != null) return _tableNameCache;
                var set = new HashSet<string>(_tableNameComparer);
                var wb = _spreadSheetDocument.WorkbookPart;
                if (wb != null)
                {
                    foreach (var ws in wb.WorksheetParts)
                    {
                        foreach (var tdp in ws.TableDefinitionParts)
                        {
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
        internal void ReserveTableName(string name)
        {
            if (string.IsNullOrWhiteSpace(name)) return;
            var cache = GetOrInitTableNameCache();
            cache.Add(name);
        }

        /// <summary>
        /// Removes the given table name from the cache. Intended for future table deletion APIs.
        /// Safe to call even if the cache hasn't been initialized.
        /// </summary>
        internal void RemoveReservedTableName(string name)
        {
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
            ExcelDocument document = new ExcelDocument();
            document.FilePath = filePath;

            // Create a spreadsheet document by supplying the filepath.
            // By default, AutoSave = true, Editable = true, and Type = xlsx.
            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook);
            document._spreadSheetDocument = spreadSheetDocument;

            // Add a WorkbookPart to the document.
            WorkbookPart workbookpart = spreadSheetDocument.AddWorkbookPart();
            workbookpart.Workbook = new Workbook();

            document._workBookPart = workbookpart;

            // Initialize document property helpers
            document.BuiltinDocumentProperties = new BuiltinDocumentProperties(document);
            document.ApplicationProperties = new ApplicationProperties(document);

            return document;
        }
        /// <summary>
        /// Loads an existing Excel document.
        /// </summary>
        /// <param name="filePath">Path to the file.</param>
        /// <param name="readOnly">Open the file in read-only mode.</param>
        /// <param name="autoSave">Enable auto-save on dispose.</param>
        /// <param name="log">Optional callback invoked when normalization failures are encountered.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Load(string filePath, bool readOnly = false, bool autoSave = false, Action<string, Exception>? log = null) {
            if (filePath == null) {
                throw new ArgumentNullException(nameof(filePath));
            }

            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }

            // Normalize content types up-front to avoid failures on malformed packages
            // where /docProps/app.xml can be incorrectly typed as application/xml.
            // Prefer in-memory normalization to avoid file-share conflicts in parallel runners.
            var bytes = File.ReadAllBytes(filePath);
            // Do NOT dispose this stream until the SpreadsheetDocument is disposed.
            // Returning a document backed by a disposed stream causes ObjectDisposedException
            // when Open XML attempts to read parts/relationships.
            var ms = new MemoryStream(bytes.Length + 4096);
            ms.Write(bytes, 0, bytes.Length);
            ms.Position = 0;
            try
            {
                Utilities.ExcelPackageUtilities.NormalizeContentTypes(ms, leaveOpen: true);
                ms.Position = 0;
                // Open from normalized memory stream to avoid touching the original file yet
                var openSettingsMem = new OpenSettings { AutoSave = autoSave };
                var memDoc = SpreadsheetDocument.Open(ms, !readOnly, openSettingsMem);
                ExcelDocument documentMem = new ExcelDocument();
                documentMem.FilePath = filePath;
                documentMem._spreadSheetDocument = memDoc;
                documentMem._workBookPart = memDoc.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
                documentMem.BuiltinDocumentProperties = new BuiltinDocumentProperties(documentMem);
                documentMem.ApplicationProperties = new ApplicationProperties(documentMem);
                return documentMem;
            }
            catch (Exception ex) when (ex is InvalidDataException || ex is OpenXmlPackageException || ex is XmlException)
            {
                ms.Dispose();
                var contextMessage = $"Failed to open '{filePath}' after normalizing package content types. The package may declare an invalid content type for '/docProps/app.xml'.";
                log?.Invoke($"{contextMessage} Inner exception: {ex.Message}", ex);
                throw new IOException($"{contextMessage} See inner exception for details.", ex);
            }
            catch
            {
                ms.Dispose();
            }
            ExcelDocument document = new ExcelDocument();
            document.FilePath = filePath;

            var openSettings = new OpenSettings {
                AutoSave = autoSave
            };

            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filePath, !readOnly, openSettings);

            document._spreadSheetDocument = spreadSheetDocument;

            //// Add a WorkbookPart to the document.
            document._workBookPart = spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");

            // Initialize document property helpers
            document.BuiltinDocumentProperties = new BuiltinDocumentProperties(document);
            document.ApplicationProperties = new ApplicationProperties(document);

            return document;
        }

        /// <summary>
        /// Validates the current spreadsheet with Open XML validator and returns error messages (if any).
        /// Useful for troubleshooting "Repaired Records" issues in Excel.
        /// </summary>
        public System.Collections.Generic.IReadOnlyList<string> ValidateOpenXml()
        {
            var list = new System.Collections.Generic.List<string>();
            if (_spreadSheetDocument == null) return list;
            // Ensure worksheet element order prior to validation so schema checks reflect final layout
            try
            {
                foreach (var sheet in Sheets)
                {
                    sheet.EnsureWorksheetElementOrder();
                }
            }
            catch { }
            var validator = new OpenXmlValidator();
            foreach (var error in validator.Validate(_spreadSheetDocument))
            {
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
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        /// <exception cref="FileNotFoundException">Thrown when the file does not exist.</exception>
        public static async Task<ExcelDocument> LoadAsync(string filePath, bool readOnly = false, bool autoSave = false) {
            if (filePath == null) {
                throw new ArgumentNullException("path");
            }
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }
            using var fileStream = new FileStream(filePath, FileMode.Open, readOnly ? FileAccess.Read : FileAccess.ReadWrite, readOnly ? FileShare.Read : FileShare.ReadWrite, 4096, FileOptions.Asynchronous);
            var memoryStream = new MemoryStream();
            await fileStream.CopyToAsync(memoryStream).ConfigureAwait(false);
            memoryStream.Seek(0, SeekOrigin.Begin);

            var openSettings = new OpenSettings {
                AutoSave = autoSave
            };

            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(memoryStream, !readOnly, openSettings);

            ExcelDocument document = new ExcelDocument {
                FilePath = filePath,
                _spreadSheetDocument = spreadSheetDocument,
                _workBookPart = spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null")
            };

            return document;
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
                string name = ValidateOrSanitizeSheetName(workSheetName, validationMode);
                ExcelSheet excelSheet = new ExcelSheet(this, _workBookPart, _spreadSheetDocument, name);
                return excelSheet;
            });
        }

        private string ValidateOrSanitizeSheetName(string name, SheetNameValidationMode mode)
        {
            // Collect existing names (case-insensitive)
            var existing = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
            foreach (var s in _workBookPart.Workbook.Sheets?.OfType<DocumentFormat.OpenXml.Spreadsheet.Sheet>() ?? System.Linq.Enumerable.Empty<DocumentFormat.OpenXml.Spreadsheet.Sheet>())
            {
                var existingName = s.Name?.Value;
                if (!string.IsNullOrEmpty(existingName)) existing.Add(existingName!);
            }

            if (mode == SheetNameValidationMode.None)
            {
                // Preserve historical behavior: default to "Sheet1" when empty
                if (string.IsNullOrEmpty(name)) name = "Sheet1";
                return name;
            }

            // Rules common to Sanitize/Strict
            static bool ContainsInvalidChars(string s)
            {
                foreach (char c in s)
                {
                    if (c == ':' || c == '\\' || c == '/' || c == '?' || c == '*' || c == '[' || c == ']') return true;
                }
                return false;
            }

            string baseName = name ?? string.Empty;
            baseName = baseName.Trim();
            baseName = baseName.Trim('\'', ' ');

            if (mode == SheetNameValidationMode.Strict)
            {
                if (string.IsNullOrEmpty(baseName)) throw new System.ArgumentException("Worksheet name cannot be empty.", nameof(name));
                if (baseName.Length > 31) throw new System.ArgumentException("Worksheet name cannot exceed 31 characters.", nameof(name));
                if (ContainsInvalidChars(baseName)) throw new System.ArgumentException("Worksheet name contains invalid characters (: \\ / ? * [ ]).", nameof(name));
                if (existing.Contains(baseName)) throw new System.ArgumentException($"Worksheet name '{baseName}' already exists.", nameof(name));
                return baseName;
            }

            // Sanitize
            var sb = new System.Text.StringBuilder(baseName.Length > 0 ? baseName.Length : 5);
            foreach (char c in baseName)
            {
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
            while (existing.Contains(candidate))
            {
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
        public void PreflightWorkbook()
        {
            foreach (var sheet in Sheets)
            {
                sheet.Preflight();
            }
        }

        /// <summary>
        /// Closes the underlying spreadsheet document.
        /// </summary>
        public void Close() {
            this._spreadSheetDocument.Dispose();
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
            // Ensure all worksheets have up-to-date dimensions and proper element ordering before saving
            foreach (var sheet in Sheets) {
                sheet.UpdateSheetDimension();
                sheet.EnsureWorksheetElementOrder();
                sheet.Commit();
            }

            // Always preflight to remove orphaned/empty containers that can trigger Excel repairs
            try { PreflightWorkbook(); } catch { }
            if (options?.SafePreflight == true)
            {
                // Already performed above; keep branch for compatibility/telemetry semantics
            }

            if (options?.SafeRepairDefinedNames == true)
            {
                try { RepairDefinedNames(save: true); } catch { }
            }
            
            _workBookPart.Workbook.Save();
            try { _spreadSheetDocument.PackageProperties.Modified = DateTime.UtcNow; } catch { }

            var path = string.IsNullOrEmpty(filePath) ? FilePath : filePath;

            // Ensure target directory is writable
            if (File.Exists(path) && new FileInfo(path).IsReadOnly) {
                throw new IOException($"Failed to save to '{path}'. The file is read-only.");
            }
            var directory = Path.GetDirectoryName(Path.GetFullPath(path));
            if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory)) {
                var dirInfo = new DirectoryInfo(directory);
                if (dirInfo.Attributes.HasFlag(FileAttributes.ReadOnly)) {
                    throw new IOException($"Failed to save to '{path}'. The directory is read-only.");
                }
            }

            // Save using OpenXML SaveAs to ensure package-level properties are persisted
            // Snapshot current package and properties
            string? pTitle = null, pCreator = null, pSubject = null, pCategory = null, pDescription = null, pKeywords = null, pLastModifiedBy = null, pVersion = null;
            DateTime? pCreated = null, pModified = null, pLastPrinted = null;
            try {
                var src = _spreadSheetDocument.PackageProperties;
                pTitle = src.Title; pCreator = src.Creator; pSubject = src.Subject; pCategory = src.Category; pDescription = src.Description;
                pKeywords = src.Keywords; pLastModifiedBy = src.LastModifiedBy; pVersion = src.Version; pCreated = src.Created; pModified = src.Modified; pLastPrinted = src.LastPrinted;
            } catch { }

            using (var fallback = new MemoryStream())
            {
                using (_spreadSheetDocument.Clone(fallback)) { }
                // Release any file handle on the original file before overwrite
                try { _spreadSheetDocument.Dispose(); } catch { }
                fallback.Position = 0;
                using (FileStream fs = new FileStream(path, FileMode.Create, FileAccess.ReadWrite, FileShare.None))
                {
                    fallback.CopyTo(fs);
                    fs.Flush();
                }
            }

            // Ensure core properties are persisted in the saved package
            try
            {
                using var pkg = Package.Open(path, FileMode.Open, FileAccess.ReadWrite);
                var dst = pkg.PackageProperties;
                dst.Title = pTitle; dst.Creator = pCreator; dst.Subject = pSubject; dst.Category = pCategory;
                dst.Description = pDescription; dst.Keywords = pKeywords; dst.LastModifiedBy = pLastModifiedBy; dst.Version = pVersion;
                dst.Created = pCreated; dst.Modified = pModified ?? DateTime.UtcNow; dst.LastPrinted = pLastPrinted;
            }
            catch { }

            try { ExcelPackageUtilities.NormalizeContentTypes(path); } catch { }
            FilePath = path;

            // Reopen as in-memory document for further operations on an expandable stream
            var fileBytes = File.ReadAllBytes(path);
            var mem = new MemoryStream(fileBytes.Length + 8192);
            mem.Write(fileBytes, 0, fileBytes.Length);
            mem.Position = 0;
            var reopenSettings = new OpenSettings { AutoSave = true };
            _spreadSheetDocument = SpreadsheetDocument.Open(mem, true, reopenSettings);
            _workBookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            _sharedStringTablePart = null;

            if (openExcel) {
                Helpers.Open(path, true);
            }

            if (options?.ValidateOpenXml == true)
            {
                var errors = ValidateOpenXml();
                if (errors.Count > 0)
                {
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
        public void SafeSave(string filePath = "", bool openExcel = false, bool writeReportOnIssues = true)
        {
            Save(filePath, openExcel);
            try
            {
                var errs = ValidateDocument();
                if (errs.Count > 0 && writeReportOnIssues)
                {
                    var target = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
                    var reportPath = System.IO.Path.ChangeExtension(target, ".xlsx.validation.txt");
                    var lines = new System.Collections.Generic.List<string>(errs.Count);
                    foreach (var e in errs)
                    {
                        lines.Add($"{e.ErrorType}: {e.Description} at {e.Path?.XPath}");
                    }
                    System.IO.File.WriteAllLines(reportPath, lines);
                }
            }
            catch { }
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
        public void Compose(string sheetName, System.Action<OfficeIMO.Excel.Fluent.SheetComposer> compose, OfficeIMO.Excel.Fluent.SheetTheme? theme = null)
        {
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
            // Ensure all worksheets have proper element ordering before saving
            foreach (var sheet in Sheets) {
                sheet.UpdateSheetDimension();
                sheet.EnsureWorksheetElementOrder();
                sheet.Commit();
            }

            // Always preflight to avoid later repair prompts
            try { PreflightWorkbook(); } catch { }
            if (options?.SafePreflight == true)
            {
                // Already performed above
            }

            if (options?.SafeRepairDefinedNames == true)
            {
                try { RepairDefinedNames(save: true); } catch { }
            }
            
            _workBookPart.Workbook.Save();

            try {
                // Serialize current document to memory snapshot
                var snapshot = new MemoryStream();
                using (_spreadSheetDocument.Clone(snapshot)) { }
                snapshot.Position = 0;

                var target = string.IsNullOrEmpty(filePath) ? FilePath : filePath;
                if (File.Exists(target) && new FileInfo(target).IsReadOnly) {
                    throw new IOException($"Failed to save to '{target}'. The file is read-only.");
                }
                var directory = Path.GetDirectoryName(Path.GetFullPath(target));
                if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory)) {
                    var dirInfo = new DirectoryInfo(directory);
                    if (dirInfo.Attributes.HasFlag(FileAttributes.ReadOnly)) {
                        throw new IOException($"Failed to save to '{target}'. The directory is read-only.");
                    }
                }

                // Snapshot props
                string? pTitle = null, pCreator = null, pSubject = null, pCategory = null, pDescription = null, pKeywords = null, pLastModifiedBy = null, pVersion = null;
                DateTime? pCreated = null, pModified = null, pLastPrinted = null;
                try { var src = _spreadSheetDocument.PackageProperties; pTitle = src.Title; pCreator = src.Creator; pSubject = src.Subject; pCategory = src.Category; pDescription = src.Description; pKeywords = src.Keywords; pLastModifiedBy = src.LastModifiedBy; pVersion = src.Version; pCreated = src.Created; pModified = src.Modified; pLastPrinted = src.LastPrinted; } catch { }

                // Release any on-disk file handle to avoid sharing violations
                try { _spreadSheetDocument.Dispose(); } catch { }

                // Write package via snapshot
                using (var fs = new FileStream(target, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 8192, FileOptions.Asynchronous)) {
                    snapshot.Position = 0;
                    await snapshot.CopyToAsync(fs, 81920, cancellationToken).ConfigureAwait(false);
                    await fs.FlushAsync(cancellationToken).ConfigureAwait(false);
                }
                // Ensure core properties persisted
                try
                {
                    using var pkg = Package.Open(target, FileMode.Open, FileAccess.ReadWrite);
                    var dst = pkg.PackageProperties;
                    dst.Title = pTitle; dst.Creator = pCreator; dst.Subject = pSubject; dst.Category = pCategory; dst.Description = pDescription; dst.Keywords = pKeywords; dst.LastModifiedBy = pLastModifiedBy; dst.Version = pVersion; dst.Created = pCreated; dst.Modified = pModified ?? DateTime.UtcNow; dst.LastPrinted = pLastPrinted;
                }
                catch { }
                try { ExcelPackageUtilities.NormalizeContentTypes(target); } catch { }
                FilePath = target;

                // Reopen as in-memory document for continued operations without locking the file
                var fileBytes = await ReadAllBytesCompatAsync(target, cancellationToken).ConfigureAwait(false);
                var mem = new MemoryStream(fileBytes.Length + 8192);
                await mem.WriteAsync(fileBytes, 0, fileBytes.Length, cancellationToken).ConfigureAwait(false);
                mem.Position = 0;
                var reopenSettings = new OpenSettings { AutoSave = true };
                _spreadSheetDocument = SpreadsheetDocument.Open(mem, true, reopenSettings);
                _workBookPart = _spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
                _sharedStringTablePart = null;
            } catch (Exception) {
                throw;
            }

            if (openExcel) {
                Open(filePath, true);
            }

            if (options?.ValidateOpenXml == true)
            {
                var errors = ValidateOpenXml();
                if (errors.Count > 0)
                {
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

            _lock?.Dispose();
            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }
}
