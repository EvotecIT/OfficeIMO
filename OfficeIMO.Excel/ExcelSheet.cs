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
                return _sheet.Name?.Value ?? string.Empty;
            }
            set {
                _sheet.Name = value;
            }
        }
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

            var workbookPart = spreadSheetDocument.WorkbookPart ??
                throw new InvalidOperationException("WorkbookPart is missing.");
            var list = workbookPart.WorksheetParts.ToList();
            foreach (var worksheetPart in list) {
                var id = workbookPart.GetIdOfPart(worksheetPart);
                if (id == _sheet.Id) {
                    _worksheetPart = worksheetPart;
                    break;
                }
            }

            if (_worksheetPart == null) {
                throw new InvalidOperationException("WorksheetPart not found for the provided sheet.");
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

            uint newId = excelDocument.id.Select(v => (uint)v).DefaultIfEmpty().Max() + 1;
            if (string.IsNullOrEmpty(name)) {
                name = "Sheet1";
            }

            // Add a WorksheetPart to the WorkbookPart.
            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
            worksheetPart.Worksheet = new Worksheet(new SheetData());

            // Add Sheets to the Workbook.
            Sheets sheets;
            var workbook = spreadSheetDocument.WorkbookPart?.Workbook ?? throw new InvalidOperationException("Workbook is missing.");
            if (workbook.Sheets != null) {
                sheets = workbook.Sheets;
            } else {
                sheets = workbook.AppendChild(new Sheets());
            }

            // Append a new worksheet and associate it with the workbook.
            Sheet sheet = new Sheet() {
                Id = spreadSheetDocument.WorkbookPart!.GetIdOfPart(worksheetPart),
                SheetId = newId,
                Name = name
            };
            sheets.Append(sheet);

            _sheet = sheet;
            Name = name;
            _worksheetPart = worksheetPart;

            excelDocument.id.Add(sheet.SheetId ?? new UInt32Value(newId));
        }

        private Cell GetCell(int row, int column) {
            if (row <= 0) {
                throw new ArgumentOutOfRangeException(nameof(row));
            }
            if (column <= 0) {
                throw new ArgumentOutOfRangeException(nameof(column));
            }

            var worksheet = _worksheetPart.Worksheet ?? (_worksheetPart.Worksheet = new Worksheet());
            SheetData sheetData = worksheet.GetFirstChild<SheetData>() ?? worksheet.AppendChild(new SheetData());

            Row? rowElement = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex != null && r.RowIndex.Value == (uint)row);
            if (rowElement == null) {
                rowElement = new Row { RowIndex = (uint)row };
                sheetData.Append(rowElement);
            }

            string cellReference = GetColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
            Cell? cell = rowElement.Elements<Cell>().FirstOrDefault(c => c.CellReference != null && c.CellReference.Value == cellReference);
            if (cell == null) {
                cell = new Cell { CellReference = cellReference };

                Cell? refCell = null;
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
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            Columns? columns = worksheet.GetFirstChild<Columns>();
            if (columns == null) {
                columns = worksheet.InsertAt(new Columns(), 0);
            }

            var font = GetDefaultFont();
            var options = new TextOptions(font);
            float zeroWidth = TextMeasurer.MeasureSize("0", options).Width;
            Dictionary<int, double> widths = new Dictionary<int, double>();

            foreach (var row in sheetData.Elements<Row>()) {
                foreach (var cell in row.Elements<Cell>()) {
                    var cellRef = cell.CellReference?.Value;
                    if (cellRef == null) continue;
                    int columnIndex = GetColumnIndex(cellRef);
                    string text = GetCellText(cell);
                    var size = TextMeasurer.MeasureSize(text, options);
                    double cellWidth = size.Width / zeroWidth + 1;
                    if (widths.ContainsKey(columnIndex)) {
                        if (cellWidth > widths[columnIndex]) widths[columnIndex] = cellWidth;
                    } else {
                        widths[columnIndex] = cellWidth;
                    }
                }
            }

            foreach (var kvp in widths) {
                Column? column = columns.Elements<Column>()
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
            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            var font = GetDefaultFont();
            var options = new TextOptions(font);
            double defaultHeight = TextMeasurer.MeasureSize("0", options).Height + 2;

            foreach (var row in sheetData.Elements<Row>()) {
                double maxHeight = 0;
                foreach (var cell in row.Elements<Cell>()) {
                    string text = GetCellText(cell);
                    var size = TextMeasurer.MeasureSize(text, options);
                    if (size.Height > maxHeight) maxHeight = size.Height;
                }
                if (maxHeight > 0) {
                    row.Height = maxHeight + 2;
                    row.CustomHeight = true;
                }
            }

            SheetFormatProperties? sheetFormat = worksheet.GetFirstChild<SheetFormatProperties>();
            if (sheetFormat == null) {
                sheetFormat = worksheet.InsertAt(new SheetFormatProperties(), 0);
            }
            sheetFormat.DefaultRowHeight = defaultHeight;
            sheetFormat.CustomHeight = true;

            worksheet.Save();
        }

        public void AddAutoFilter(string range, Dictionary<uint, IEnumerable<string>>? filterCriteria = null) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            Worksheet worksheet = _worksheetPart.Worksheet;

            AutoFilter? existing = worksheet.Elements<AutoFilter>().FirstOrDefault();
            if (existing != null) {
                worksheet.RemoveChild(existing);
            }

            AutoFilter autoFilter = new AutoFilter { Reference = range };

            if (filterCriteria != null) {
                foreach (KeyValuePair<uint, IEnumerable<string>> criteria in filterCriteria) {
                    FilterColumn filterColumn = new FilterColumn { ColumnId = criteria.Key };
                    Filters filters = new Filters();
                    foreach (string value in criteria.Value) {
                        filters.Append(new Filter { Val = value });
                    }

                    filterColumn.Append(filters);
                    autoFilter.Append(filterColumn);
                }
            }

            worksheet.Append(autoFilter);
            worksheet.Save();
        }

        public void AddTable(string range, bool hasHeader, string name, TableStyle style) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            var cells = range.Split(':');
            if (cells.Length != 2) {
                throw new ArgumentException("Invalid range format", nameof(range));
            }

            string startRef = cells[0];
            string endRef = cells[1];

            int startColumnIndex = GetColumnIndex(startRef);
            int endColumnIndex = GetColumnIndex(endRef);

            uint columnsCount = (uint)(endColumnIndex - startColumnIndex + 1);

            var tableDefinitionPart = _worksheetPart.AddNewPart<TableDefinitionPart>();
            uint tableId = (uint)(_worksheetPart.TableDefinitionParts.Count() + 1);

            if (string.IsNullOrEmpty(name)) {
                name = $"Table{tableId}";
            }

            var table = new Table {
                Id = tableId,
                Name = name,
                DisplayName = name,
                Reference = range,
                HeaderRowCount = hasHeader ? (uint)1 : (uint)0,
                TotalsRowShown = false
            };

            var tableColumns = new TableColumns { Count = columnsCount };
            for (uint i = 0; i < columnsCount; i++) {
                tableColumns.Append(new TableColumn { Id = i + 1, Name = $"Column{i + 1}" });
            }
            table.Append(tableColumns);

            table.Append(new TableStyleInfo {
                Name = style.ToString(),
                ShowFirstColumn = false,
                ShowLastColumn = false,
                ShowRowStripes = true,
                ShowColumnStripes = false
            });

            tableDefinitionPart.Table = table;

            TableParts? tableParts = _worksheetPart.Worksheet.Elements<TableParts>().FirstOrDefault();
            if (tableParts == null) {
                tableParts = new TableParts { Count = 1 };
                _worksheetPart.Worksheet.Append(tableParts);
            } else {
                tableParts.Count = (tableParts.Count ?? 0) + 1;
            }

            var relId = _worksheetPart.GetIdOfPart(tableDefinitionPart);
            tableParts.Append(new TablePart { Id = relId });

            _worksheetPart.Worksheet.Save();
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

        public void SetCellValue(int row, int column, DateTimeOffset value, bool autoFitColumns = false, bool autoFitRows = false) {
            SetCellValue(row, column, value.UtcDateTime, autoFitColumns, autoFitRows);
        }

        public void SetCellValue(int row, int column, TimeSpan value, bool autoFitColumns = false, bool autoFitRows = false) {
            Cell cell = GetCell(row, column);
            cell.CellValue = new CellValue(value.TotalDays.ToString(CultureInfo.InvariantCulture));
            cell.DataType = CellValues.Number;
            if (autoFitColumns) AutoFitColumns();
            if (autoFitRows) AutoFitRows();
        }

        public void SetCellValue(int row, int column, uint value, bool autoFitColumns = false, bool autoFitRows = false) {
            SetCellValue(row, column, (double)value, autoFitColumns, autoFitRows);
        }

        public void SetCellValue(int row, int column, ulong value, bool autoFitColumns = false, bool autoFitRows = false) {
            SetCellValue(row, column, (double)value, autoFitColumns, autoFitRows);
        }

        public void SetCellValue(int row, int column, ushort value, bool autoFitColumns = false, bool autoFitRows = false) {
            SetCellValue(row, column, (double)value, autoFitColumns, autoFitRows);
        }

        public void SetCellValue(int row, int column, byte value, bool autoFitColumns = false, bool autoFitRows = false) {
            SetCellValue(row, column, (double)value, autoFitColumns, autoFitRows);
        }

        public void SetCellValue(int row, int column, sbyte value, bool autoFitColumns = false, bool autoFitRows = false) {
            SetCellValue(row, column, (double)value, autoFitColumns, autoFitRows);
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

        public void SetCellFormat(int row, int column, string numberFormat) {
            Cell cell = GetCell(row, column);

            var workbookPart = _excelDocument._spreadSheetDocument.WorkbookPart ??
                throw new InvalidOperationException("WorkbookPart is missing.");
            WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();

            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();

            stylesheet.Fonts ??= new Fonts(new DocumentFormat.OpenXml.Spreadsheet.Font());
            stylesheet.Fonts.Count = (uint)stylesheet.Fonts.Count();

            stylesheet.Fills ??= new Fills(new Fill());
            stylesheet.Fills.Count = (uint)stylesheet.Fills.Count();

            stylesheet.Borders ??= new Borders(new Border());
            stylesheet.Borders.Count = (uint)stylesheet.Borders.Count();

            stylesheet.CellStyleFormats ??= new CellStyleFormats(new CellFormat());
            stylesheet.CellStyleFormats.Count = (uint)stylesheet.CellStyleFormats.Count();

            stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            if (stylesheet.CellFormats.Count == null || stylesheet.CellFormats.Count.Value == 0) {
                stylesheet.CellFormats.Count = 1;
            }

            stylesheet.NumberingFormats ??= new NumberingFormats();

            NumberingFormat? existingFormat = stylesheet.NumberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(n => n.FormatCode != null && n.FormatCode.Value == numberFormat);

            uint numberFormatId;
            if (existingFormat?.NumberFormatId != null) {
                numberFormatId = existingFormat.NumberFormatId.Value;
            } else {
                numberFormatId = stylesheet.NumberingFormats.Elements<NumberingFormat>().Any()
                    ? stylesheet.NumberingFormats.Elements<NumberingFormat>().Max(n => n.NumberFormatId?.Value ?? 0) + 1
                    : 164U;
                NumberingFormat numberingFormat = new NumberingFormat {
                    NumberFormatId = numberFormatId,
                    FormatCode = StringValue.FromString(numberFormat)
                };
                stylesheet.NumberingFormats.Append(numberingFormat);
                stylesheet.NumberingFormats.Count = (uint)stylesheet.NumberingFormats.Count();
            }

            var cellFormats = stylesheet.CellFormats.Elements<CellFormat>().ToList();
            int formatIndex = cellFormats.FindIndex(cf => cf.NumberFormatId != null && cf.NumberFormatId.Value == numberFormatId && cf.ApplyNumberFormat != null && cf.ApplyNumberFormat.Value);
            if (formatIndex == -1) {
                CellFormat cellFormat = new CellFormat {
                    NumberFormatId = numberFormatId,
                    ApplyNumberFormat = true
                };
                stylesheet.CellFormats.Append(cellFormat);
                stylesheet.CellFormats.Count = (uint)stylesheet.CellFormats.Count();
                formatIndex = cellFormats.Count;
            }

            cell.StyleIndex = (uint)formatIndex;
            stylesPart.Stylesheet.Save();
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
                case DateTimeOffset dto:
                    SetCellValue(row, column, dto, autoFitColumns, autoFitRows);
                    break;
                case TimeSpan ts:
                    SetCellValue(row, column, ts, autoFitColumns, autoFitRows);
                    break;
                case bool b:
                    SetCellValue(row, column, b, autoFitColumns, autoFitRows);
                    break;
                case uint ui:
                    SetCellValue(row, column, ui, autoFitColumns, autoFitRows);
                    break;
                case ulong ul:
                    SetCellValue(row, column, ul, autoFitColumns, autoFitRows);
                    break;
                case ushort us:
                    SetCellValue(row, column, us, autoFitColumns, autoFitRows);
                    break;
                case byte by:
                    SetCellValue(row, column, by, autoFitColumns, autoFitRows);
                    break;
                case sbyte sb:
                    SetCellValue(row, column, sb, autoFitColumns, autoFitRows);
                    break;
                case short sh:
                    SetCellValue(row, column, (double)sh, autoFitColumns, autoFitRows);
                    break;
                default:
                    string text = value?.ToString() ?? string.Empty;
                    SetCellValue(row, column, text, autoFitColumns, autoFitRows);
                    break;
            }
        }

        public void SetCellValue<T>(int row, int column, T? value, bool autoFitColumns = false, bool autoFitRows = false) where T : struct {
            if (value.HasValue) {
                SetCellValue(row, column, value.Value, autoFitColumns, autoFitRows);
            } else {
                SetCellValue(row, column, string.Empty, autoFitColumns, autoFitRows);
            }
        }
    }
}
