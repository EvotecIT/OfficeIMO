using System;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Collections.Generic;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using SixLabors.Fonts;
using SixLaborsColor = SixLabors.ImageSharp.Color;

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
        private readonly ReaderWriterLockSlim _lock = new ReaderWriterLockSlim(LockRecursionPolicy.SupportsRecursion);

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

        private void WriteLock(Action action) {
            _lock.EnterWriteLock();
            try {
                action();
            } finally {
                _lock.ExitWriteLock();
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

        /// <summary>
        /// Automatically fits all columns based on their content.
        /// </summary>
        public void AutoFitColumns() {
            WriteLock(() => {
                var worksheet = _worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) return;

                var columns = worksheet.GetFirstChild<Columns>();
                HashSet<int> columnIndexes = new HashSet<int>();

                foreach (var row in sheetData.Elements<Row>()) {
                    foreach (var cell in row.Elements<Cell>()) {
                        if (cell.CellReference == null) continue;
                        columnIndexes.Add(GetColumnIndex(cell.CellReference.Value));
                    }
                }

                if (columns != null) {
                    foreach (var column in columns.Elements<Column>()) {
                        uint min = column.Min?.Value ?? 0;
                        uint max = column.Max?.Value ?? 0;
                        for (uint i = min; i <= max; i++) {
                            columnIndexes.Add((int)i);
                        }
                    }
                }

                foreach (int index in columnIndexes.OrderBy(i => i)) {
                    AutoFitColumn(index);
                }

                worksheet.Save();
            });
        }

        /// <summary>
        /// Automatically fits all rows based on their content.
        /// </summary>
        public void AutoFitRows() {
            WriteLock(() => {
                var worksheet = _worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) return;

                var rowIndexes = sheetData.Elements<Row>()
                    .Select(r => (int)r.RowIndex!.Value)
                    .ToList();

                foreach (int rowIndex in rowIndexes) {
                    AutoFitRow(rowIndex);
                }

                var sheetFormat = worksheet.GetFirstChild<SheetFormatProperties>();
                bool anyCustom = sheetData.Elements<Row>()
                    .Any(r => r.CustomHeight != null && r.CustomHeight.Value);

                if (anyCustom) {
                    if (sheetFormat == null) {
                        sheetFormat = worksheet.InsertAt(new SheetFormatProperties(), 0);
                    }
                    sheetFormat.DefaultRowHeight = 15;
                    sheetFormat.CustomHeight = true;
                } else if (sheetFormat != null) {
                    sheetFormat.Remove();
                }

                worksheet.Save();
            });
        }

        public void AutoFitColumn(int columnIndex) {
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
            double width = 0;

            foreach (var row in sheetData.Elements<Row>()) {
                var cell = row.Elements<Cell>()
                    .FirstOrDefault(c => c.CellReference != null && GetColumnIndex(c.CellReference.Value) == columnIndex);
                if (cell == null) continue;
                string text = GetCellText(cell);
                if (string.IsNullOrWhiteSpace(text)) continue;
                var size = TextMeasurer.MeasureSize(text, options);
                double cellWidth = size.Width / zeroWidth + 1;
                if (cellWidth > width) width = cellWidth;
            }

            Column column = columns.Elements<Column>()
                .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);
            if (width > 0) {
                if (column == null) {
                    column = new Column { Min = (uint)columnIndex, Max = (uint)columnIndex };
                    columns.Append(column);
                }
                column.Width = width;
                column.CustomWidth = true;
                column.BestFit = true;
            } else if (column != null) {
                column.Remove();
            }

            worksheet.Save();
        }

        public void AutoFitRow(int rowIndex) {
            var worksheet = _worksheetPart.Worksheet;
            SheetData sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) return;

            var font = GetDefaultFont();
            var options = new TextOptions(font);

            double defaultHeight = 15;
            double ToPoints(double height) {
                return height * 72.0 / options.Dpi;
            }

            Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == (uint)rowIndex);
            if (row == null) return;

            double maxHeight = 0;
            foreach (var cell in row.Elements<Cell>()) {
                string text = GetCellText(cell);
                if (string.IsNullOrWhiteSpace(text)) continue;
                var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                double lineHeight = lines.Max(line => ToPoints(TextMeasurer.MeasureSize(line, options).Height));
                double cellHeight = lineHeight * lines.Length;
                if (cellHeight > maxHeight) maxHeight = cellHeight;
            }

            if (maxHeight > 0) {
                row.Height = maxHeight + 2;
                row.CustomHeight = true;
                var sheetFormat = worksheet.GetFirstChild<SheetFormatProperties>();
                if (sheetFormat == null) {
                    sheetFormat = worksheet.InsertAt(new SheetFormatProperties(), 0);
                }
                sheetFormat.DefaultRowHeight = defaultHeight;
                sheetFormat.CustomHeight = true;
            } else {
                row.Height = null;
                row.CustomHeight = null;
            }

            worksheet.Save();
        }

        public void AddAutoFilter(string range, Dictionary<uint, IEnumerable<string>> filterCriteria = null) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            Worksheet worksheet = _worksheetPart.Worksheet;

            AutoFilter existing = worksheet.Elements<AutoFilter>().FirstOrDefault();
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

            var tableParts = _worksheetPart.Worksheet.Elements<TableParts>().FirstOrDefault();
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

        public void AddConditionalRule(string range, ConditionalFormattingOperatorValues @operator, string formula1, string? formula2 = null) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            Worksheet worksheet = _worksheetPart.Worksheet;

            int priority = 1;
            var existingRules = worksheet.Descendants<ConditionalFormattingRule>();
            if (existingRules.Any()) {
                priority = existingRules.Max(r => r.Priority?.Value ?? 0) + 1;
            }

            ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
            };

            ConditionalFormattingRule rule = new ConditionalFormattingRule {
                Type = ConditionalFormatValues.CellIs,
                Operator = @operator,
                Priority = priority
            };

            rule.Append(new Formula(formula1));
            if (formula2 != null) {
                rule.Append(new Formula(formula2));
            }

            conditionalFormatting.Append(rule);
            worksheet.Append(conditionalFormatting);
            worksheet.Save();
        }

        private static string ConvertColor(SixLaborsColor color) {
            return "FF" + color.ToHexColor();
        }

        public void AddConditionalColorScale(string range, SixLaborsColor startColor, SixLaborsColor endColor) {
            AddConditionalColorScale(range, ConvertColor(startColor), ConvertColor(endColor));
        }

        public void AddConditionalColorScale(string range, string startColor, string endColor) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            Worksheet worksheet = _worksheetPart.Worksheet;

            int priority = 1;
            var existingRules = worksheet.Descendants<ConditionalFormattingRule>();
            if (existingRules.Any()) {
                priority = existingRules.Max(r => r.Priority?.Value ?? 0) + 1;
            }

            ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
            };

            ConditionalFormattingRule rule = new ConditionalFormattingRule {
                Type = ConditionalFormatValues.ColorScale,
                Priority = priority
            };

            ColorScale colorScale = new ColorScale();
            colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
            colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
            colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = startColor });
            colorScale.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = endColor });
            rule.Append(colorScale);

            conditionalFormatting.Append(rule);
            worksheet.Append(conditionalFormatting);
            worksheet.Save();
        }

        public void AddConditionalDataBar(string range, SixLaborsColor color) {
            AddConditionalDataBar(range, ConvertColor(color));
        }

        public void AddConditionalDataBar(string range, string color) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            Worksheet worksheet = _worksheetPart.Worksheet;

            int priority = 1;
            var existingRules = worksheet.Descendants<ConditionalFormattingRule>();
            if (existingRules.Any()) {
                priority = existingRules.Max(r => r.Priority?.Value ?? 0) + 1;
            }

            ConditionalFormatting conditionalFormatting = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
            };

            ConditionalFormattingRule rule = new ConditionalFormattingRule {
                Type = ConditionalFormatValues.DataBar,
                Priority = priority
            };

            DataBar dataBar = new DataBar { ShowValue = true };
            dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
            dataBar.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
            dataBar.Append(new DocumentFormat.OpenXml.Spreadsheet.Color { Rgb = color });
            rule.Append(dataBar);

            conditionalFormatting.Append(rule);
            worksheet.Append(conditionalFormatting);
            worksheet.Save();
        }

        public void SetCellValue(int row, int column, string value) {
            WriteLock(() => {
                Cell cell = GetCell(row, column);
                int sharedStringIndex = _excelDocument.GetSharedStringIndex(value);
                cell.CellValue = new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.SharedString;
            });
        }

        public void SetCellValue(int row, int column, double value) {
            WriteLock(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        public void SetCellValue(int row, int column, decimal value) {
            WriteLock(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        public void SetCellValue(int row, int column, DateTime value) {
            WriteLock(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.ToOADate().ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        public void SetCellValue(int row, int column, DateTimeOffset value) {
            SetCellValue(row, column, value.UtcDateTime);
        }

        public void SetCellValue(int row, int column, TimeSpan value) {
            WriteLock(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.TotalDays.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        public void SetCellValue(int row, int column, uint value) {
            SetCellValue(row, column, (double)value);
        }

        public void SetCellValue(int row, int column, ulong value) {
            SetCellValue(row, column, (double)value);
        }

        public void SetCellValue(int row, int column, ushort value) {
            SetCellValue(row, column, (double)value);
        }

        public void SetCellValue(int row, int column, byte value) {
            SetCellValue(row, column, (double)value);
        }

        public void SetCellValue(int row, int column, sbyte value) {
            SetCellValue(row, column, (double)value);
        }

        public void SetCellValue(int row, int column, bool value) {
            WriteLock(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value ? "1" : "0");
                cell.DataType = CellValues.Boolean;
            });
        }

        public void SetCellFormula(int row, int column, string formula) {
            WriteLock(() => {
                Cell cell = GetCell(row, column);
                cell.CellFormula = new CellFormula(formula);
            });
        }

        public void SetCellFormat(int row, int column, string numberFormat) {
            Cell cell = GetCell(row, column);

            WorkbookStylesPart stylesPart = _excelDocument._spreadSheetDocument.WorkbookPart.WorkbookStylesPart;
            if (stylesPart == null) {
                stylesPart = _excelDocument._spreadSheetDocument.WorkbookPart.AddNewPart<WorkbookStylesPart>();
            }

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

            NumberingFormat existingFormat = stylesheet.NumberingFormats.Elements<NumberingFormat>()
                .FirstOrDefault(n => n.FormatCode != null && n.FormatCode.Value == numberFormat);

            uint numberFormatId;
            if (existingFormat != null) {
                numberFormatId = existingFormat.NumberFormatId.Value;
            } else {
                numberFormatId = stylesheet.NumberingFormats.Elements<NumberingFormat>().Any()
                    ? stylesheet.NumberingFormats.Elements<NumberingFormat>().Max(n => n.NumberFormatId.Value) + 1
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

        public void SetCellValue(int row, int column, object value) {
            WriteLock(() => {
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
                    case DateTimeOffset dto:
                        SetCellValue(row, column, dto);
                        break;
                    case TimeSpan ts:
                        SetCellValue(row, column, ts);
                        break;
                    case bool b:
                        SetCellValue(row, column, b);
                        break;
                    case uint ui:
                        SetCellValue(row, column, ui);
                        break;
                    case ulong ul:
                        SetCellValue(row, column, ul);
                        break;
                    case ushort us:
                        SetCellValue(row, column, us);
                        break;
                    case byte by:
                        SetCellValue(row, column, by);
                        break;
                    case sbyte sb:
                        SetCellValue(row, column, sb);
                        break;
                    case short sh:
                        SetCellValue(row, column, (double)sh);
                        break;
                    default:
                        if (value != null) {
                            SetCellValue(row, column, value.ToString());
                        } else {
                            SetCellValue(row, column, string.Empty);
                        }
                        break;
                }
            });
        }

        public void SetCellValue<T>(int row, int column, T? value) where T : struct {
            if (value.HasValue) {
                SetCellValue(row, column, value.Value);
            } else {
                SetCellValue(row, column, string.Empty);
            }
        }
    }
}
