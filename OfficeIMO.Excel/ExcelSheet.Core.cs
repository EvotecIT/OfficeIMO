using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using System.Globalization;
using System.Threading;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a single worksheet within an <see cref="ExcelDocument"/>.
    /// </summary>
    public partial class ExcelSheet : IDisposable {
        private readonly Sheet _sheet;
        internal Sheet SheetElement => _sheet;

        /// <summary>
        /// Gets or sets the worksheet name.
        /// </summary>
        public string Name {
            get {
                return _sheet.Name?.Value ?? string.Empty;
            }
            set {
                _excelDocument.RenameWorkSheet(this, value, SheetNameValidationMode.Strict);
            }
        }
        private readonly UInt32Value _id;
        private readonly WorksheetPart _worksheetPart;
        internal WorksheetPart WorksheetPart => _worksheetPart;
        private readonly SpreadsheetDocument _spreadSheetDocument;
        private readonly ExcelDocument _excelDocument;
        private bool _isBatchOperation = false;
        private bool _hasWorksheetMutations;
        private bool _requiresSavePreparation;
        private readonly List<TableDefinitionPart> _pendingTableDefinitionPartSaves = new();
        private readonly object _batchLock = new object();
        private Row? _lastAccessedRow;
        private int _lastAccessedRowIndex;
        private Cell? _lastAccessedCell;
        private int _lastAccessedCellRowIndex;
        private int _lastAccessedCellColumnIndex;
        private SheetData? _sheetDataCache;
        private static int _instancesCreated;

        internal static int InstancesCreatedForTests => Volatile.Read(ref _instancesCreated);

        internal static void ResetInstanceCountForTests() => Interlocked.Exchange(ref _instancesCreated, 0);

        /// <summary>
        /// Override execution policy for this sheet. Null = inherit from document.
        /// </summary>
        public ExecutionPolicy? ExecutionOverride { get; set; }

        /// <summary>
        /// Gets the effective execution policy for this sheet.
        /// </summary>
        internal ExecutionPolicy EffectiveExecution => ExecutionOverride ?? _excelDocument.Execution;

        /// <summary>
        /// Begin a no-lock context where operations bypass locking.
        /// </summary>
        public NoLockContext BeginNoLock() => new();

        /// <summary>
        /// Represents a scope where worksheet operations bypass locking.
        /// </summary>
        public sealed class NoLockContext : IDisposable {
            private readonly IDisposable _scope;
            internal NoLockContext() => _scope = Locking.EnterNoLockScope();

            /// <summary>
            /// Ends the no-lock scope and restores normal locking behavior.
            /// </summary>
            public void Dispose() => _scope.Dispose();
        }

        /// <summary>
        /// Returns the used range of this worksheet as an A1 string by leveraging the read bridge.
        /// </summary>
        public string GetUsedRangeA1() {
            using var reader = _excelDocument.CreateReader();
            var sh = reader.GetSheet(this.Name);
            return sh.GetUsedRangeA1();
        }

        /// <summary>
        /// Initializes a worksheet from an existing <see cref="Sheet"/> element.
        /// </summary>
        /// <param name="excelDocument">Parent document.</param>
        /// <param name="spreadSheetDocument">Open XML spreadsheet document.</param>
        /// <param name="sheet">Underlying sheet element.</param>
        public ExcelSheet(ExcelDocument excelDocument, SpreadsheetDocument spreadSheetDocument, Sheet sheet) {
            _excelDocument = excelDocument;
            _sheet = sheet;
            _spreadSheetDocument = spreadSheetDocument;
            _hasWorksheetMutations = excelDocument.IsPackageDirty;
            _requiresSavePreparation = excelDocument.IsPackageDirty;

            var workbookPart = spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            _worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            _id = sheet.SheetId!;
            Interlocked.Increment(ref _instancesCreated);
        }

        /// <summary>
        /// Creates a new worksheet and appends it to the workbook.
        /// </summary>
        /// <param name="excelDocument">Parent document.</param>
        /// <param name="workbookpart">Workbook part to add the worksheet to.</param>
        /// <param name="spreadSheetDocument">Open XML spreadsheet document.</param>
        /// <param name="name">Worksheet name.</param>
        public ExcelSheet(ExcelDocument excelDocument, WorkbookPart workbookpart, SpreadsheetDocument spreadSheetDocument, string name) {
            _excelDocument = excelDocument;
            _spreadSheetDocument = spreadSheetDocument;
            _hasWorksheetMutations = true;
            _requiresSavePreparation = true;

            UInt32Value id = excelDocument.id.Max(v => v.Value) + 1;
            if (name == "") {
                name = "Sheet1";
            }

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            var spWorkbookPart = spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            var workbook = spWorkbookPart.Workbook ??= new Workbook();
            Sheets sheets;
            if (workbook.Sheets != null) {
                sheets = workbook.Sheets;
            } else {
                sheets = workbook.AppendChild(new Sheets());
            }

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() {
                Id = spreadSheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = id,
                Name = name
            };
            sheets.Append(sheet);

            this._sheet = sheet;
            this._id = sheet.SheetId!;
            this._worksheetPart = worksheetPart;

            excelDocument.id.Add(id);
            Interlocked.Increment(ref _instancesCreated);
        }

        private Cell GetCell(int row, int column) {
            if (row <= 0) {
                throw new ArgumentOutOfRangeException(nameof(row));
            }
            if (column <= 0) {
                throw new ArgumentOutOfRangeException(nameof(column));
            }

            SheetData sheetData = GetOrCreateSheetData();

            Row? rowElement = null;
            Row? insertAfterRow = null;
            bool createdRowElement = false;
            if (_lastAccessedRow != null && ReferenceEquals(_lastAccessedRow.Parent, sheetData)) {
                if (_lastAccessedRowIndex == row) {
                    rowElement = _lastAccessedRow;
                } else if (_lastAccessedRowIndex < row) {
                    insertAfterRow = _lastAccessedRow;
                    for (Row? next = _lastAccessedRow.NextSibling<Row>(); next != null; next = next.NextSibling<Row>()) {
                        if (next.RowIndex == null) {
                            continue;
                        }

                        int nextRowIndex = (int)next.RowIndex.Value;
                        if (nextRowIndex == row) {
                            rowElement = next;
                            break;
                        }

                        if (nextRowIndex > row) {
                            break;
                        }

                        insertAfterRow = next;
                    }
                }
            }

            if (rowElement == null && insertAfterRow == null) {
                foreach (Row r in sheetData.Elements<Row>()) {
                    if (r.RowIndex != null) {
                        if (r.RowIndex.Value == (uint)row) {
                            rowElement = r;
                            break;
                        }
                        if (r.RowIndex.Value < (uint)row) {
                            insertAfterRow = r;
                        } else {
                            break;
                        }
                    }
                }
            }

            if (rowElement == null) {
                rowElement = new Row { RowIndex = (uint)row };
                createdRowElement = true;
                if (insertAfterRow != null) {
                    sheetData.InsertAfter(rowElement, insertAfterRow);
                } else {
                    // Insert at beginning
                    var firstRow = sheetData.Elements<Row>().FirstOrDefault();
                    if (firstRow != null) {
                        sheetData.InsertBefore(rowElement, firstRow);
                    } else {
                        sheetData.Append(rowElement);
                    }
                }
            }

            if (createdRowElement) {
                Cell createdCell = new Cell { CellReference = A1.CellReference(row, column) };
                rowElement.Append(createdCell);
                _lastAccessedRow = rowElement;
                _lastAccessedRowIndex = row;
                _lastAccessedCell = createdCell;
                _lastAccessedCellRowIndex = row;
                _lastAccessedCellColumnIndex = column;
                return createdCell;
            }

            // Find or create cell with proper ordering (by numeric column index)
            Cell? cell = null;
            Cell? insertAfterCell = null;
            int targetColumnIndex = column;

            if (_lastAccessedCell != null
                && _lastAccessedCellRowIndex == row
                && ReferenceEquals(_lastAccessedCell.Parent, rowElement)) {
                if (_lastAccessedCellColumnIndex == targetColumnIndex) {
                    cell = _lastAccessedCell;
                } else if (_lastAccessedCellColumnIndex < targetColumnIndex) {
                    insertAfterCell = _lastAccessedCell;
                    for (Cell? next = _lastAccessedCell.NextSibling<Cell>(); next != null; next = next.NextSibling<Cell>()) {
                        if (next.CellReference?.Value is not string nextRefValue || nextRefValue.Length == 0) {
                            continue;
                        }

                        int nextColumnIndex = GetColumnIndex(nextRefValue);
                        if (nextColumnIndex == targetColumnIndex) {
                            cell = next;
                            break;
                        }

                        if (nextColumnIndex > targetColumnIndex) {
                            break;
                        }

                        insertAfterCell = next;
                    }
                }
            }

            if (cell == null && insertAfterCell == null) {
                foreach (Cell c in rowElement.Elements<Cell>()) {
                    if (c.CellReference?.Value is not string existingRefValue || existingRefValue.Length == 0) {
                        continue;
                    }

                    int existingColumnIndex = GetColumnIndex(existingRefValue);
                    if (existingColumnIndex == targetColumnIndex) {
                        cell = c;
                        break;
                    }
                    if (existingColumnIndex < targetColumnIndex) {
                        insertAfterCell = c;
                        continue;
                    }
                    // existingColumnIndex > targetColumnIndex => insert before this cell
                    break;
                }
            }

            if (cell == null) {
                cell = new Cell { CellReference = A1.CellReference(row, column) };
                if (insertAfterCell != null) {
                    rowElement.InsertAfter(cell, insertAfterCell);
                } else {
                    // Insert at beginning or append when row is empty or existing first cell has larger column index
                    var firstCell = rowElement.Elements<Cell>().FirstOrDefault();
                    if (firstCell != null) {
                        if (firstCell.CellReference?.Value is string firstRefValue && firstRefValue.Length > 0) {
                            if (GetColumnIndex(firstRefValue) > targetColumnIndex) {
                                rowElement.InsertBefore(cell, firstCell);
                            } else {
                                rowElement.Append(cell);
                            }
                        } else {
                            rowElement.Append(cell);
                        }
                    } else {
                        rowElement.Append(cell);
                    }
                }
            }

            _lastAccessedRow = rowElement;
            _lastAccessedRowIndex = row;
            _lastAccessedCell = cell;
            _lastAccessedCellRowIndex = row;
            _lastAccessedCellColumnIndex = column;
            return cell;
        }

        private SheetData GetOrCreateSheetData() {
            var worksheet = WorksheetRoot;
            if (_sheetDataCache != null && ReferenceEquals(_sheetDataCache.Parent, worksheet)) {
                return _sheetDataCache;
            }

            _sheetDataCache = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());
            return _sheetDataCache;
        }

        private Row GetOrCreateRowElement(SheetData sheetData, int rowIndex) {
            foreach (Row row in sheetData.Elements<Row>()) {
                if (row.RowIndex != null) {
                    if (row.RowIndex.Value == (uint)rowIndex) {
                        return row;
                    }
                    if (row.RowIndex.Value > (uint)rowIndex) {
                        var inserted = new Row { RowIndex = (uint)rowIndex };
                        sheetData.InsertBefore(inserted, row);
                        return inserted;
                    }
                }
            }

            var appended = new Row { RowIndex = (uint)rowIndex };
            sheetData.Append(appended);
            return appended;
        }

        private static string GetColumnName(int columnIndex) {
            return A1.ColumnIndexToLetters(columnIndex);
        }

        private static int GetColumnIndex(string cellReference) {
            int columnIndex = 0;
            for (int i = 0; i < cellReference.Length; i++) {
                char ch = cellReference[i];
                if (ch >= 'A' && ch <= 'Z') {
                    columnIndex = (columnIndex * 26) + (ch - 'A' + 1);
                    continue;
                }

                if (ch >= 'a' && ch <= 'z') {
                    columnIndex = (columnIndex * 26) + (ch - 'a' + 1);
                    continue;
                }

                if (columnIndex > 0) {
                    break;
                }
            }
            return columnIndex;
        }

        private static int GetRowIndex(string cellReference) {
            int rowIndex = 0;
            for (int i = 0; i < cellReference.Length; i++) {
                char ch = cellReference[i];
                if (ch >= '0' && ch <= '9') {
                    rowIndex = (rowIndex * 10) + (ch - '0');
                }
            }

            return rowIndex;
        }

        // Exposed as internal so other components in the same assembly (e.g., SheetComposer)
        // can reuse the shared-string/inline-string resolution logic when they already have
        // a Cell reference. Prefer using public TryGetCellText(row,col, out text) when possible.
        internal string GetCellText(Cell cell) {
            // Shared string lookup
            if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString) {
                var raw = cell.CellValue?.InnerText;
                if (!string.IsNullOrEmpty(raw) && TryParseCellTextSharedStringIndex(raw, out int id)) {
                    return BuildCellTextSharedStringSnapshot().Get(id) ?? string.Empty;
                }

                return string.Empty;
            }

            // Inline string
            if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString) {
                var inline = cell.InlineString;
                if (inline != null) {
                    if (inline.Text != null) {
                        return inline.Text.Text ?? string.Empty;
                    }
                    var sb = new StringBuilder();
                    foreach (var r in inline.Elements<Run>()) {
                        if (r.Text != null) sb.Append(r.Text.Text);
                    }
                    return sb.ToString();
                }
                return string.Empty;
            }

            // Default: take cell value as-is (numbers, booleans, etc.)
            return cell.CellValue?.InnerText ?? string.Empty;
        }

        private SharedStringCache BuildCellTextSharedStringSnapshot() {
            return SharedStringCache.Build(_spreadSheetDocument);
        }

        private void ClearCellTextSharedStringCache() {
        }

        private static bool TryParseCellTextSharedStringIndex(string? text, out int index) {
            index = 0;
            if (string.IsNullOrEmpty(text)) {
                return false;
            }

            int parsed = 0;
            for (int i = 0; i < text!.Length; i++) {
                int digit = text[i] - '0';
                if ((uint)digit > 9U || parsed > (int.MaxValue - digit) / 10) {
                    return int.TryParse(text, NumberStyles.Integer, CultureInfo.InvariantCulture, out index);
                }

                parsed = (parsed * 10) + digit;
            }

            index = parsed;
            return true;
        }

        private void WriteLock(Action action) {
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), () => {
                action();
                MarkRequiresSavePreparation();
            });
        }

        private void WriteLockConditional(Action action) {
            // If we're already in a batch operation or in a NoLock scope,
            // just execute the action directly
            if (_isBatchOperation || Locking.IsNoLock) {
                action();
                MarkRequiresSavePreparation();
            } else {
                _excelDocument.MaterializeDeferredDataSetImport();
                WriteLock(action);
            }
        }

        private OfficeFontInfo? GetWorkbookDefaultFontInfo() {
            try {
                var workbookPart = WorkbookPartRoot;
                var stylesPart = workbookPart?.WorkbookStylesPart;
                var stylesheet = stylesPart?.Stylesheet;
                var fonts = stylesheet?.Fonts;
                var firstFont = fonts?.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().FirstOrDefault();
                if (firstFont == null) return null;

                var fontName = firstFont.GetFirstChild<FontName>()?.Val?.Value;
                var fontSize = firstFont.GetFirstChild<FontSize>()?.Val?.Value ?? 11.0;
                bool bold = firstFont.GetFirstChild<Bold>() != null;
                bool italic = firstFont.GetFirstChild<Italic>() != null;
                bool underline = firstFont.GetFirstChild<Underline>() != null;

                if (!string.IsNullOrEmpty(fontName)) {
                    return new OfficeFontInfo(fontName, fontSize, GetOfficeFontStyle(bold, italic, underline));
                }
            } catch {
                // ignore
            }
            return null;
        }

        private OfficeFontInfo GetCellFontInfo(Cell cell, OfficeFontInfo fallbackFontInfo) {
            if (cell.StyleIndex == null) return fallbackFontInfo;

            var workbookPart = WorkbookPartRoot;
            var stylesPart = workbookPart?.WorkbookStylesPart;
            var stylesheet = stylesPart?.Stylesheet;
            var fonts = stylesheet?.Fonts;
            var cellFormats = stylesheet?.CellFormats;
            if (fonts == null || cellFormats == null) return fallbackFontInfo;

            var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)cell.StyleIndex.Value);
            if (cellFormat?.FontId == null) return fallbackFontInfo;

            var fontElement = fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAtOrDefault((int)cellFormat.FontId.Value);
            if (fontElement == null) return fallbackFontInfo;

            return CreateFontInfoFromOpenXml(fontElement, (float)fallbackFontInfo.Size);
        }

        private static OfficeFontInfo CreateFontInfoFromOpenXml(DocumentFormat.OpenXml.Spreadsheet.Font fontElement, float fallbackSize) {
            var fontName = fontElement.GetFirstChild<FontName>()?.Val?.Value;
            var fontSize = fontElement.GetFirstChild<FontSize>()?.Val?.Value ?? fallbackSize;
            bool bold = fontElement.GetFirstChild<Bold>() != null;
            bool italic = fontElement.GetFirstChild<Italic>() != null;
            bool underline = fontElement.GetFirstChild<Underline>() != null;

            return new OfficeFontInfo(fontName, fontSize, GetOfficeFontStyle(bold, italic, underline));
        }

        private static OfficeFontStyle GetOfficeFontStyle(bool bold, bool italic, bool underline) {
            var style = OfficeFontStyle.Regular;
            if (bold) style |= OfficeFontStyle.Bold;
            if (italic) style |= OfficeFontStyle.Italic;
            if (underline) style |= OfficeFontStyle.Underline;
            return style;
        }

        /// <summary>
        /// Releases resources held by this worksheet.
        /// </summary>
        public void Dispose() {
            // No local lock to dispose anymore - using document's lock
        }

        /// <summary>
        /// Persists pending changes on this worksheet to its underlying OpenXml part.
        /// </summary>
        internal void Commit() {
            if (_pendingTableDefinitionPartSaves.Count > 0) {
                foreach (var tableDefinitionPart in _pendingTableDefinitionPartSaves) {
                    tableDefinitionPart.Table?.Save();
                }

                _pendingTableDefinitionPartSaves.Clear();
            }

            _worksheetPart?.Worksheet?.Save();
            _requiresSavePreparation = false;
        }

        internal bool RequiresSavePreparation => _requiresSavePreparation;

        internal void MarkRequiresSavePreparation() {
            _requiresSavePreparation = true;
            _excelDocument.MarkRequiresSavePreflight();
        }

        internal void DeferTableDefinitionPartSave(TableDefinitionPart tableDefinitionPart) {
            if (!_pendingTableDefinitionPartSaves.Contains(tableDefinitionPart)) {
                _pendingTableDefinitionPartSaves.Add(tableDefinitionPart);
            }

            MarkRequiresSavePreparation();
        }
    }
}
