using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents an Excel document and provides methods for creating,
    /// loading and saving spreadsheets.
    /// </summary>
    public partial class ExcelDocument : IDisposable, IAsyncDisposable {
        private readonly ReaderWriterLockSlim _lock = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);
        internal List<UInt32Value> id = new List<UInt32Value>() { 0 };

        /// <summary>
        /// Gets a list of worksheets contained in the document.
        /// </summary>
        public List<ExcelSheet> Sheets {
            get {
                _lock.EnterUpgradeableReadLock();
                try {
                    List<ExcelSheet> listExcel = new List<ExcelSheet>();
                    List<Sheet>? elements = null;
                    if (_spreadSheetDocument.WorkbookPart!.Workbook!.Sheets != null) {
                        elements = _spreadSheetDocument.WorkbookPart.Workbook.Sheets.OfType<Sheet>().ToList();
                        foreach (Sheet s in elements) {
                            listExcel.Add(new ExcelSheet(this, _spreadSheetDocument, s));
                        }
                    }

                    _lock.EnterWriteLock();
                    try {
                        id.Clear();
                        id.Add(0);
                        if (elements != null) {
                            foreach (Sheet s in elements) {
                                if (s.SheetId != null && !id.Contains(s.SheetId)) {
                                    id.Add(s.SheetId);
                                }
                            }
                        }
                    } finally {
                        _lock.ExitWriteLock();
                    }

                    return listExcel;
                } finally {
                    _lock.ExitUpgradeableReadLock();
                }
            }
        }

        /// <summary>
        /// Underlying Open XML spreadsheet document instance.
        /// </summary>
        public SpreadsheetDocument _spreadSheetDocument = null!;
        private WorkbookPart _workBookPart = null!;
        private SharedStringTablePart _sharedStringTablePart = null!;

        /// <summary>
        /// Path to the file backing this document.
        /// </summary>
        public string FilePath = string.Empty;

        /// <summary>
        /// FileOpenAccess of the document
        /// </summary>
        public FileAccess FileOpenAccess => _spreadSheetDocument.FileOpenAccess;

        internal SharedStringTablePart SharedStringTablePart {
            get {
                _lock.EnterUpgradeableReadLock();
                try {
                    if (_sharedStringTablePart == null) {
                        _lock.EnterWriteLock();
                        try {
                            if (_workBookPart.GetPartsOfType<SharedStringTablePart>().Any()) {
                                _sharedStringTablePart = _workBookPart.GetPartsOfType<SharedStringTablePart>().First();
                            } else {
                                _sharedStringTablePart = _workBookPart.AddNewPart<SharedStringTablePart>();
                                _sharedStringTablePart.SharedStringTable = new SharedStringTable();
                            }
                        } finally {
                            _lock.ExitWriteLock();
                        }
                    }

                    return _sharedStringTablePart;
                } finally {
                    _lock.ExitUpgradeableReadLock();
                }
            }
        }

        internal int GetSharedStringIndex(string text) {
            _lock.EnterUpgradeableReadLock();
            try {
                var sharedStringTable = SharedStringTablePart.SharedStringTable;
                int index = 0;
                foreach (SharedStringItem item in sharedStringTable.Elements<SharedStringItem>()) {
                    if (item.InnerText == text) {
                        return index;
                    }
                    index++;
                }

                _lock.EnterWriteLock();
                try {
                    sharedStringTable.AppendChild(new SharedStringItem(new Text(text)));
                    sharedStringTable.Save();
                } finally {
                    _lock.ExitWriteLock();
                }

                return index;
            } finally {
                _lock.ExitUpgradeableReadLock();
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

            return document;
        }
        /// <summary>
        /// Loads an existing Excel document.
        /// </summary>
        /// <param name="filePath">Path to the file.</param>
        /// <param name="readOnly">Open the file in read-only mode.</param>
        /// <param name="autoSave">Enable auto-save on dispose.</param>
        /// <returns>Loaded <see cref="ExcelDocument"/> instance.</returns>
        public static ExcelDocument Load(string filePath, bool readOnly = false, bool autoSave = false) {
            if (filePath == null) {
                throw new ArgumentNullException(nameof(filePath));
            }

            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }
            ExcelDocument document = new ExcelDocument();
            document.FilePath = filePath;

            var openSettings = new OpenSettings {
                AutoSave = autoSave
            };

            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filePath, !readOnly, openSettings);

            document._spreadSheetDocument = spreadSheetDocument;

            //// Add a WorkbookPart to the document.
            document._workBookPart = document._spreadSheetDocument.WorkbookPart;

            return document;
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
                throw new ArgumentNullException(nameof(filePath));
            }
            if (!File.Exists(filePath)) {
                throw new FileNotFoundException($"File '{filePath}' doesn't exist.", filePath);
            }
            using var fileStream = new FileStream(filePath, FileMode.Open, readOnly ? FileAccess.Read : FileAccess.ReadWrite, readOnly ? FileShare.Read : FileShare.ReadWrite, 4096, FileOptions.Asynchronous);
            var memoryStream = new MemoryStream();
            await fileStream.CopyToAsync(memoryStream);
            memoryStream.Seek(0, SeekOrigin.Begin);

            var openSettings = new OpenSettings {
                AutoSave = autoSave
            };

            SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(memoryStream, !readOnly, openSettings);

            ExcelDocument document = new ExcelDocument {
                FilePath = filePath,
                _spreadSheetDocument = spreadSheetDocument,
                _workBookPart = spreadSheetDocument.WorkbookPart
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
            excelDocument.AddWorkSheet(workSheetName);
            return excelDocument;
        }

        /// <summary>
        /// Adds a worksheet to the document.
        /// </summary>
        /// <param name="workSheetName">Worksheet name.</param>
        /// <returns>Created <see cref="ExcelSheet"/> instance.</returns>
        public ExcelSheet AddWorkSheet(string workSheetName = "") {
            _lock.EnterWriteLock();
            try {
                ExcelSheet excelSheet = new ExcelSheet(this, _workBookPart, _spreadSheetDocument, workSheetName);

                return excelSheet;
            } finally {
                _lock.ExitWriteLock();
            }
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
            this._workBookPart.Workbook.Save();

            this.Open(filePath, openExcel);
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
        /// Asynchronously saves the document.
        /// </summary>
        /// <param name="filePath">Optional path to save to.</param>
        /// <param name="openExcel">Whether to open Excel after saving.</param>
        /// <param name="cancellationToken">Cancellation token.</param>
        /// <returns>A task representing the asynchronous operation.</returns>
        public async Task SaveAsync(string filePath, bool openExcel, CancellationToken cancellationToken = default) {
            _workBookPart.Workbook.Save();

            try {
                if (!string.IsNullOrEmpty(filePath)) {
                    if (File.Exists(filePath) && new FileInfo(filePath).IsReadOnly) {
                        throw new IOException($"Failed to save to '{filePath}'. The file is read-only.");
                    }

                    var directory = Path.GetDirectoryName(Path.GetFullPath(filePath));
                    if (!string.IsNullOrEmpty(directory) && Directory.Exists(directory)) {
                        var dirInfo = new DirectoryInfo(directory);
                        if (dirInfo.Attributes.HasFlag(FileAttributes.ReadOnly)) {
                            throw new IOException($"Failed to save to '{filePath}'. The directory is read-only.");
                        }
                    }

                    FileStream fs;
                    try {
                        fs = new FileStream(filePath, FileMode.Create, FileAccess.ReadWrite, FileShare.None, 4096, FileOptions.Asynchronous);
                    } catch (UnauthorizedAccessException ex) {
                        throw new IOException($"Failed to save to '{filePath}'. Access denied or path is read-only.", ex);
                    }

                    using (fs) {
                        using (var clone = _spreadSheetDocument.Clone(fs)) {
                        }
                        await fs.FlushAsync(cancellationToken);
                    }
                    FilePath = filePath;
                }
            } catch (Exception ex) when (ex is IOException || ex is UnauthorizedAccessException) {
                throw;
            }

            if (openExcel) {
                Open(filePath, true);
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

        /// <summary>
        /// Releases resources used by the document.
        /// </summary>
        private bool _disposed;

        public void Dispose() {
            DisposeAsync().GetAwaiter().GetResult();
        }

        public async ValueTask DisposeAsync() {
            if (_disposed) {
                return;
            }

            if (this._spreadSheetDocument != null) {
                try {
                    if (this._spreadSheetDocument.AutoSave && this._spreadSheetDocument.FileOpenAccess != FileAccess.Read) {
                        _workBookPart?.Workbook.Save();
                    }

                    await Task.Run(() => this._spreadSheetDocument.Dispose());
                } catch {
                    // ignored
                }
                this._spreadSheetDocument = null;
            }

            _disposed = true;
            GC.SuppressFinalize(this);
        }
    }
}
