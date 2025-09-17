using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.Fonts;
using System;
using System.Collections;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Diagnostics;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;
using SixLaborsColor = SixLabors.ImageSharp.Color;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a single worksheet within an <see cref="ExcelDocument"/>.
    /// </summary>
    public partial class ExcelSheet : IDisposable {
        private readonly Sheet _sheet;

        /// <summary>
        /// Gets or sets the worksheet name.
        /// </summary>
        public string Name {
            get {
                return _sheet.Name?.Value ?? string.Empty;
            }
            set {
                _sheet.Name = value;
            }
        }
        private readonly UInt32Value _id;
        private readonly WorksheetPart _worksheetPart;
        private readonly SpreadsheetDocument _spreadSheetDocument;
        private readonly ExcelDocument _excelDocument;
        private bool _isBatchOperation = false;
        private readonly object _batchLock = new object();
        private static int _nextTableId = 1;
        private static readonly object _tableIdLock = new object();

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
        public sealed class NoLockContext : IDisposable
        {
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
        public string GetUsedRangeA1()
        {
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

            var workbookPart = spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            _worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            _id = sheet.SheetId!;
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

            UInt32Value id = excelDocument.id.Max(v => v.Value) + 1;
            if (name == "") {
                name = "Sheet1";
            }

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            var spWorkbookPart = spreadSheetDocument.WorkbookPart ?? throw new InvalidOperationException("WorkbookPart is null");
            Sheets sheets;
            if (spWorkbookPart.Workbook.Sheets != null) {
                sheets = spWorkbookPart.Workbook.Sheets;
            } else {
                sheets = spWorkbookPart.Workbook.AppendChild(new Sheets());
            }

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() {
                Id = spreadSheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                SheetId = id,
                Name = name
            };
            sheets.Append(sheet);

            this._sheet = sheet;
            this.Name = name;
            this._id = sheet.SheetId!;
            this._worksheetPart = worksheetPart;

            excelDocument.id.Add(id);
        }

        private Cell GetCell(int row, int column) {
            if (row <= 0) {
                throw new ArgumentOutOfRangeException(nameof(row));
            }
            if (column <= 0) {
                throw new ArgumentOutOfRangeException(nameof(column));
            }

              SheetData? sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
              if (sheetData == null) {
                  sheetData = _worksheetPart.Worksheet.AppendChild(new SheetData());
              }

            // Find or create row with proper ordering
              Row? rowElement = null;
              Row? insertAfterRow = null;
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

            if (rowElement == null) {
                rowElement = new Row { RowIndex = (uint)row };
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

            string cellReference = GetColumnName(column) + row.ToString(CultureInfo.InvariantCulture);

            // Find or create cell with proper ordering (by numeric column index)
            Cell? cell = null;
            Cell? insertAfterCell = null;
            int targetColumnIndex = column;
            foreach (Cell c in rowElement.Elements<Cell>()) {
                if (c.CellReference?.Value is { Length: > 0 } existingRef) {
                    int existingColumnIndex = GetColumnIndex(existingRef);
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
                cell = new Cell { CellReference = cellReference };
                if (insertAfterCell != null) {
                    rowElement.InsertAfter(cell, insertAfterCell);
                } else {
                    // Insert at beginning or append when row is empty or existing first cell has larger column index
                    var firstCell = rowElement.Elements<Cell>().FirstOrDefault();
                    if (firstCell != null) {
                        if (firstCell.CellReference?.Value is { Length: > 0 } firstRef && GetColumnIndex(firstRef) > targetColumnIndex) {
                            rowElement.InsertBefore(cell, firstCell);
                        } else {
                            rowElement.Append(cell);
                        }
                    } else {
                        rowElement.Append(cell);
                    }
                }
            }

            return cell;
        }

        private static string GetColumnName(int columnIndex) {
            int dividend = columnIndex;
            StringBuilder columnName = new StringBuilder();

            while (dividend > 0) {
                int modulo = (dividend - 1) % 26;
                columnName.Insert(0, Convert.ToChar(65 + modulo));
                dividend = (dividend - modulo) / 26;
            }

            return columnName.ToString();
        }

        private static int GetColumnIndex(string cellReference) {
            ArgumentNullException.ThrowIfNull(cellReference);
            int columnIndex = 0;
            foreach (char ch in cellReference.Where(char.IsLetter)) {
                columnIndex = (columnIndex * 26) + (ch - 'A' + 1);
            }
            return columnIndex;
        }

        private static int GetRowIndex(string cellReference) {
            var digits = new string(cellReference.Where(char.IsDigit).ToArray());
            return int.Parse(digits, CultureInfo.InvariantCulture);
        }

        // Exposed as internal so other components in the same assembly (e.g., SheetComposer)
        // can reuse the shared-string/inline-string resolution logic when they already have
        // a Cell reference. Prefer using public TryGetCellText(row,col, out text) when possible.
        internal string GetCellText(Cell cell) {
            // Shared string lookup
            if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.SharedString)
            {
                var raw = cell.CellValue?.InnerText;
                if (!string.IsNullOrEmpty(raw) && int.TryParse(raw, out int id))
                {
                    var sst = _excelDocument.SharedStringTablePart?.SharedStringTable;
                    if (sst != null)
                    {
                        var item = sst.Elements<SharedStringItem>().ElementAtOrDefault(id);
                        if (item != null)
                        {
                            // Prefer direct Text element when present; otherwise concatenate run texts
                            if (item.Text != null)
                            {
                                return item.Text.Text ?? string.Empty;
                            }
                            var sb = new StringBuilder();
                            foreach (var t in item.Descendants<Text>())
                            {
                                sb.Append(t.Text);
                            }
                            return sb.ToString();
                        }
                    }
                }
                return string.Empty;
            }

            // Inline string
            if (cell.DataType?.Value == DocumentFormat.OpenXml.Spreadsheet.CellValues.InlineString)
            {
                var inline = cell.InlineString;
                if (inline != null)
                {
                    if (inline.Text != null)
                    {
                        return inline.Text.Text ?? string.Empty;
                    }
                    var sb = new StringBuilder();
                    foreach (var r in inline.Elements<Run>())
                    {
                        if (r.Text != null) sb.Append(r.Text.Text);
                    }
                    return sb.ToString();
                }
                return string.Empty;
            }

            // Default: take cell value as-is (numbers, booleans, etc.)
            return cell.CellValue?.InnerText ?? string.Empty;
        }

        private void WriteLock(Action action) {
            Locking.ExecuteWrite(_excelDocument.EnsureLock(), action);
        }

        private void WriteLockConditional(Action action) {
            // If we're already in a batch operation or in a NoLock scope,
            // just execute the action directly
            if (_isBatchOperation || Locking.IsNoLock) {
                action();
            } else {
                WriteLock(action);
            }
        }

        private SixLabors.Fonts.Font GetDefaultFont() {
            // Try to use the workbook's default font if present
            var wf = GetWorkbookDefaultFont();
            if (wf != null) return wf;

            string[] preferred = { "Calibri", "Arial", "Liberation Sans", "DejaVu Sans", "Times New Roman" };

            foreach (var name in preferred) {
                try {
                    var font = SystemFonts.CreateFont(name, 11);
                    if (IsFontUsable(font)) return font;
                } catch (FontFamilyNotFoundException) {
                    // Try next option
                }
            }

            foreach (var family in SystemFonts.Collection.Families) {
                try {
                    var font = family.CreateFont(11);
                    if (IsFontUsable(font)) return font;
                } catch {
                    // Skip fonts that cannot be loaded or measured
                }
            }

            // Fallback to first available family without validation
            return SystemFonts.Collection.Families.First().CreateFont(11);
        }

        private SixLabors.Fonts.Font? GetWorkbookDefaultFont() {
            try {
                var workbookPart = _spreadSheetDocument.WorkbookPart;
                var stylesPart = workbookPart?.WorkbookStylesPart;
                var stylesheet = stylesPart?.Stylesheet;
                var fonts = stylesheet?.Fonts;
                var firstFont = fonts?.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().FirstOrDefault();
                if (firstFont == null) return null;

                var fontName = firstFont.GetFirstChild<FontName>()?.Val?.Value;
                var fontSize = firstFont.GetFirstChild<FontSize>()?.Val?.Value ?? 11.0;
                bool bold = firstFont.GetFirstChild<Bold>() != null;
                bool italic = firstFont.GetFirstChild<Italic>() != null;

                var style = bold && italic ? FontStyle.BoldItalic : bold ? FontStyle.Bold : italic ? FontStyle.Italic : FontStyle.Regular;
                if (!string.IsNullOrEmpty(fontName)) {
                    try {
                        return SystemFonts.CreateFont(fontName!, (float)fontSize, style);
                    } catch (FontFamilyNotFoundException) {
                        return null;
                    }
                }
            } catch {
                // ignore
            }
            return null;
        }

        private static bool IsFontUsable(SixLabors.Fonts.Font font) {
            try {
                TextMeasurer.MeasureSize("0", new TextOptions(font));
                return true;
            } catch {
                return false;
            }
        }

        private SixLabors.Fonts.Font GetCellFont(Cell cell) {
            var defaultFont = GetDefaultFont();
            if (cell.StyleIndex == null) return defaultFont;

            var workbookPart = _spreadSheetDocument.WorkbookPart;
            var stylesPart = workbookPart?.WorkbookStylesPart;
            var stylesheet = stylesPart?.Stylesheet;
            var fonts = stylesheet?.Fonts;
            var cellFormats = stylesheet?.CellFormats;
            if (fonts == null || cellFormats == null) return defaultFont;

            var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault((int)cell.StyleIndex.Value);
            if (cellFormat?.FontId == null) return defaultFont;

            var fontElement = fonts.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>().ElementAtOrDefault((int)cellFormat.FontId.Value);
            if (fontElement == null) return defaultFont;

            var fontName = fontElement.GetFirstChild<FontName>()?.Val?.Value;
            var fontSize = fontElement.GetFirstChild<FontSize>()?.Val?.Value ?? defaultFont.Size;
            bool bold = fontElement.GetFirstChild<Bold>() != null;
            bool italic = fontElement.GetFirstChild<Italic>() != null;

            try {
                var style = bold && italic ? FontStyle.BoldItalic : bold ? FontStyle.Bold : italic ? FontStyle.Italic : FontStyle.Regular;
                if (!string.IsNullOrEmpty(fontName)) {
                    return SystemFonts.CreateFont(fontName!, (float)fontSize, style);
                }
                return defaultFont.Family.CreateFont((float)fontSize, style);
            } catch (FontFamilyNotFoundException) {
                var fallbackStyle = bold && italic ? FontStyle.BoldItalic : bold ? FontStyle.Bold : italic ? FontStyle.Italic : FontStyle.Regular;
                return defaultFont.Family.CreateFont((float)fontSize, fallbackStyle);
            }
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
            _worksheetPart?.Worksheet?.Save();
        }
    }
}
