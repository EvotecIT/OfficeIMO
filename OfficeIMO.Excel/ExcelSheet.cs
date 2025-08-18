using System;
using System.Globalization;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

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

        public void SetCellValue(int row, int column, string value) {
            Cell cell = GetCell(row, column);
            int sharedStringIndex = _excelDocument.GetSharedStringIndex(value);
            cell.CellValue = new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.SharedString;
        }

        public void SetCellValue(int row, int column, double value) {
            Cell cell = GetCell(row, column);
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
        }

        public void SetCellValue(int row, int column, decimal value) {
            Cell cell = GetCell(row, column);
            cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
        }

        public void SetCellValue(int row, int column, DateTime value) {
            Cell cell = GetCell(row, column);
            cell.CellValue = new CellValue(value.ToOADate().ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
        }

        public void SetCellValue(int row, int column, bool value) {
            Cell cell = GetCell(row, column);
            cell.CellValue = new CellValue(value ? "1" : "0");
            cell.DataType = CellValues.Boolean;
        }

        public void SetCellFormula(int row, int column, string formula) {
            Cell cell = GetCell(row, column);
            cell.CellFormula = new CellFormula(formula);
        }

        public void SetCellValue(int row, int column, object value) {
            switch (value) {
                case string s:
                    SetCellValue(row, column, s);
                    break;
                case double d:
                    SetCellValue(row, column, d);
                    break;
                case float f:
                    SetCellValue(row, column, Convert.ToDouble(f));
                    break;
                case decimal dec:
                    SetCellValue(row, column, dec);
                    break;
                case int i:
                    SetCellValue(row, column, (double)i);
                    break;
                case long l:
                    SetCellValue(row, column, (double)l);
                    break;
                case DateTime dt:
                    SetCellValue(row, column, dt);
                    break;
                case bool b:
                    SetCellValue(row, column, b);
                    break;
                default:
                    if (value != null) {
                        SetCellValue(row, column, value.ToString());
                    } else {
                        SetCellValue(row, column, string.Empty);
                    }
                    break;
            }
        }
    }
}
