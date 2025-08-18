using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.Fonts;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a single worksheet within an <see cref="ExcelDocument"/>.
    /// </summary>
    public class ExcelSheet {
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

            Row rowElement = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)row);
            if (rowElement == null) {
                rowElement = new Row { RowIndex = (uint)row };
                sheetData.Append(rowElement);
            }

            string cellReference = GetColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
            Cell cell = rowElement.Elements<Cell>().FirstOrDefault(c => c.CellReference != null && c.CellReference.Value == cellReference);
            if (cell == null) {
                cell = new Cell { CellReference = cellReference };

                Cell refCell = null;
                foreach (Cell c in rowElement.Elements<Cell>()) {
                    if (string.Compare(c.CellReference?.Value, cellReference, StringComparison.Ordinal) > 0) {
                        refCell = c;
                        break;
                    }
                }
                if (refCell != null) {
                    rowElement.InsertBefore(cell, refCell);
                } else {
                    rowElement.Append(cell);
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

        public void AutoFitColumns() {
            var worksheet = _worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            var columns = worksheet.GetFirstChild<Columns>();
            if (columns == null) {
                columns = worksheet.InsertAt(new Columns(), 0);
            }

            var font = GetDefaultFont();
            var options = new TextOptions(font);
            float zeroWidth = TextMeasurer.MeasureSize("0", options).Width;
            Dictionary<int, double> widths = new Dictionary<int, double>();

            foreach (var row in sheetData.Elements<Row>()) {
                foreach (var cell in row.Elements<Cell>()) {
                    if (cell.CellReference == null) continue;
                    int columnIndex = GetColumnIndex(cell.CellReference.Value);
                    string text = GetCellText(cell);
                    var size = TextMeasurer.MeasureSize(text ?? string.Empty, options);
                    double cellWidth = size.Width / zeroWidth + 1;
                    if (widths.ContainsKey(columnIndex)) {
                        if (cellWidth > widths[columnIndex]) widths[columnIndex] = cellWidth;
                    } else {
                        widths[columnIndex] = cellWidth;
                    }
                }
            }

            foreach (var kvp in widths) {
                Column column = columns.Elements<Column>()
                    .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)kvp.Key && c.Max.Value >= (uint)kvp.Key);
                if (column == null) {
                    column = new Column { Min = (uint)kvp.Key, Max = (uint)kvp.Key };
                    columns.Append(column);
                }
                column.Width = kvp.Value;
                column.CustomWidth = true;
                column.BestFit = true;
            }

            worksheet.Save();
        }

        public void AutoFitRows() {
            var worksheet = _worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            var font = GetDefaultFont();
            var options = new TextOptions(font);
            double defaultHeight = TextMeasurer.MeasureSize("0", options).Height + 2;

            foreach (var row in sheetData.Elements<Row>()) {
                double maxHeight = 0;
                foreach (var cell in row.Elements<Cell>()) {
                    string text = GetCellText(cell);
                    var size = TextMeasurer.MeasureSize(text ?? string.Empty, options);
                    if (size.Height > maxHeight) maxHeight = size.Height;
                }
                if (maxHeight > 0) {
                    row.Height = maxHeight + 2;
                    row.CustomHeight = true;
                }
            }

            var sheetFormat = worksheet.GetFirstChild<SheetFormatProperties>();
            if (sheetFormat == null) {
                sheetFormat = worksheet.InsertAt(new SheetFormatProperties(), 0);
            }
            sheetFormat.DefaultRowHeight = defaultHeight;
            sheetFormat.CustomHeight = true;

            worksheet.Save();
        }

        public void SetCellValue(int row, int column, string value, bool autoFitColumns = false, bool autoFitRows = false) {
            Cell cell = GetCell(row, column);
            int sharedStringIndex = _excelDocument.GetSharedStringIndex(value);
            cell.CellValue = new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.SharedString;
            if (autoFitColumns) AutoFitColumns();
            if (autoFitRows) AutoFitRows();
        }

        public void SetCellValue(int row, int column, double value, bool autoFitColumns = false, bool autoFitRows = false) {
            Cell cell = GetCell(row, column);
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
            if (autoFitColumns) AutoFitColumns();
            if (autoFitRows) AutoFitRows();
        }

        public void SetCellValue(int row, int column, decimal value, bool autoFitColumns = false, bool autoFitRows = false) {
            Cell cell = GetCell(row, column);
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
            if (autoFitColumns) AutoFitColumns();
            if (autoFitRows) AutoFitRows();
        }

        public void SetCellValue(int row, int column, DateTime value, bool autoFitColumns = false, bool autoFitRows = false) {
            Cell cell = GetCell(row, column);
            cell.CellValue = new CellValue(value.ToOADate().ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
            if (autoFitColumns) AutoFitColumns();
            if (autoFitRows) AutoFitRows();
        }

        public void SetCellValue(int row, int column, bool value, bool autoFitColumns = false, bool autoFitRows = false) {
            Cell cell = GetCell(row, column);
            cell.CellValue = new CellValue(value ? "1" : "0");
            cell.DataType = CellValues.Boolean;
            if (autoFitColumns) AutoFitColumns();
            if (autoFitRows) AutoFitRows();
        }

        public void SetCellFormula(int row, int column, string formula, bool autoFitColumns = false, bool autoFitRows = false) {
            Cell cell = GetCell(row, column);
            cell.CellFormula = new CellFormula(formula);
            if (autoFitColumns) AutoFitColumns();
            if (autoFitRows) AutoFitRows();
        }

        public void SetCellValue(int row, int column, object value, bool autoFitColumns = false, bool autoFitRows = false) {
            switch (value) {
                case string s:
                    SetCellValue(row, column, s, autoFitColumns, autoFitRows);
                    break;
                case double d:
                    SetCellValue(row, column, d, autoFitColumns, autoFitRows);
                    break;
                case float f:
                    SetCellValue(row, column, Convert.ToDouble(f), autoFitColumns, autoFitRows);
                    break;
                case decimal dec:
                    SetCellValue(row, column, dec, autoFitColumns, autoFitRows);
                    break;
                case int i:
                    SetCellValue(row, column, (double)i, autoFitColumns, autoFitRows);
                    break;
                case long l:
                    SetCellValue(row, column, (double)l, autoFitColumns, autoFitRows);
                    break;
                case DateTime dt:
                    SetCellValue(row, column, dt, autoFitColumns, autoFitRows);
                    break;
                case bool b:
                    SetCellValue(row, column, b, autoFitColumns, autoFitRows);
                    break;
                default:
                    if (value != null) {
                        SetCellValue(row, column, value.ToString(), autoFitColumns, autoFitRows);
                    } else {
                        SetCellValue(row, column, string.Empty, autoFitColumns, autoFitRows);
                    }
                    break;
            }
        }
    }
}
