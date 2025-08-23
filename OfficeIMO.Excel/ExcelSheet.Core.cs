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
                return _sheet.Name;
            }
            set {
                _sheet.Name = value;
            }
        }
        private readonly UInt32Value Id;
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

        public sealed class NoLockContext : IDisposable
        {
            private readonly IDisposable _scope;
            internal NoLockContext() => _scope = Locking.EnterNoLockScope();
            public void Dispose() => _scope.Dispose();
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

            var list = _spreadSheetDocument.WorkbookPart.WorksheetParts.ToList();
            foreach (var worksheetPart in list) {
                var id = spreadSheetDocument.WorkbookPart.GetIdOfPart(worksheetPart);
                if (id == _sheet.Id) {
                    _worksheetPart = worksheetPart;
                }
            }
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

            UInt32Value id = excelDocument.id.Max() + 1;
            if (name == "") {
                name = "Sheet1";
            }

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets = null;
            if (spreadSheetDocument.WorkbookPart.Workbook.Sheets != null) {
                sheets = spreadSheetDocument.WorkbookPart.Workbook.Sheets;
            } else {
                sheets = spreadSheetDocument.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());
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
            this.Id = sheet.SheetId;
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

            SheetData sheetData = _worksheetPart.Worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                sheetData = _worksheetPart.Worksheet.AppendChild(new SheetData());
            }

            // Find or create row with proper ordering
            Row rowElement = null;
            Row insertAfterRow = null;
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

            // Find or create cell with proper ordering
            Cell cell = null;
            Cell insertAfterCell = null;
            foreach (Cell c in rowElement.Elements<Cell>()) {
                if (c.CellReference != null) {
                    if (c.CellReference.Value == cellReference) {
                        cell = c;
                        break;
                    }
                    var compareResult = string.Compare(c.CellReference.Value, cellReference, StringComparison.Ordinal);
                    if (compareResult < 0) {
                        insertAfterCell = c;
                    } else {
                        break;
                    }
                }
            }

            if (cell == null) {
                cell = new Cell { CellReference = cellReference };
                if (insertAfterCell != null) {
                    rowElement.InsertAfter(cell, insertAfterCell);
                } else {
                    // Insert at beginning
                    var firstCell = rowElement.Elements<Cell>().FirstOrDefault();
                    if (firstCell != null && string.Compare(firstCell.CellReference?.Value, cellReference, StringComparison.Ordinal) > 0) {
                        rowElement.InsertBefore(cell, firstCell);
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

        private string GetCellText(Cell cell) {
            if (cell.CellValue == null) return string.Empty;
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString) {
                if (int.TryParse(value, out int id)) {
                    var item = _excelDocument.SharedStringTablePart.SharedStringTable.Elements<SharedStringItem>().ElementAt(id);
                    return item.InnerText;
                }
            }
            return value;
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

        private static SixLabors.Fonts.Font GetDefaultFont() {
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

            var stylesPart = _spreadSheetDocument.WorkbookPart.WorkbookStylesPart;
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

            try {
                if (!string.IsNullOrEmpty(fontName)) {
                    return SystemFonts.CreateFont(fontName, (float)fontSize, bold ? FontStyle.Bold : FontStyle.Regular);
                }
                return defaultFont.Family.CreateFont((float)fontSize, bold ? FontStyle.Bold : FontStyle.Regular);
            } catch (FontFamilyNotFoundException) {
                return defaultFont.Family.CreateFont((float)fontSize, bold ? FontStyle.Bold : FontStyle.Regular);
            }
        }

        public void Dispose() {
            // No local lock to dispose anymore - using document's lock
        }
    }
}
