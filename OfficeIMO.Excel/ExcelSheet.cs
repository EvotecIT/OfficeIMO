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
    public class ExcelSheet : IDisposable {
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
        private readonly AsyncLocal<bool> _skipWriteLock = new AsyncLocal<bool>();

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
            _lock.EnterWriteLock();
            try {
                action();
            } finally {
                _lock.ExitWriteLock();
            }
        }

        private void WriteLockConditional(Action action) {
            if (_skipWriteLock.Value) {
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
            WriteLock(() => {
                var worksheet = _worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) return;

                var columns = worksheet.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = worksheet.InsertAt(new Columns(), 0);
                }

                double width = 0;

                foreach (var row in sheetData.Elements<Row>()) {
                    var cell = row.Elements<Cell>()
                        .FirstOrDefault(c => c.CellReference != null && GetColumnIndex(c.CellReference.Value) == columnIndex);
                    if (cell == null) continue;
                    string text = GetCellText(cell);
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    var font = GetCellFont(cell);
                    var options = new TextOptions(font);
                    float zeroWidth = TextMeasurer.MeasureSize("0", options).Width;
                    var size = TextMeasurer.MeasureSize(text, options);
                    double cellWidth = zeroWidth == 0 ? 0 : size.Width / zeroWidth + 1;
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
            });
        }

        public void SetColumnWidth(int columnIndex, double width) {
            WriteLock(() => {
                var worksheet = _worksheetPart.Worksheet;
                var columns = worksheet.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = worksheet.InsertAt(new Columns(), 0);
                }
                var column = columns.Elements<Column>()
                    .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);
                if (column == null) {
                    column = new Column { Min = (uint)columnIndex, Max = (uint)columnIndex };
                    columns.Append(column);
                }
                column.Width = width;
                column.CustomWidth = true;
                worksheet.Save();
            });
        }

        public void SetColumnHidden(int columnIndex, bool hidden) {
            WriteLock(() => {
                var worksheet = _worksheetPart.Worksheet;
                var columns = worksheet.GetFirstChild<Columns>();
                if (columns == null) {
                    columns = worksheet.InsertAt(new Columns(), 0);
                }
                var column = columns.Elements<Column>()
                    .FirstOrDefault(c => c.Min != null && c.Max != null && c.Min.Value <= (uint)columnIndex && c.Max.Value >= (uint)columnIndex);
                if (column == null) {
                    column = new Column { Min = (uint)columnIndex, Max = (uint)columnIndex };
                    columns.Append(column);
                }
                column.Hidden = hidden ? true : (bool?)null;
                worksheet.Save();
            });
        }

        public void AutoFitRow(int rowIndex) {
            WriteLock(() => {
                var worksheet = _worksheetPart.Worksheet;
                SheetData sheetData = worksheet.GetFirstChild<SheetData>();
                if (sheetData == null) return;

                Row row = sheetData.Elements<Row>().FirstOrDefault(r => r.RowIndex == (uint)rowIndex);
                if (row == null) return;

                const double defaultHeight = 15;
                const double pointsPerInch = 72.0;

                double maxHeight = 0;
                foreach (var cell in row.Elements<Cell>()) {
                    string text = GetCellText(cell);
                    if (string.IsNullOrWhiteSpace(text)) continue;
                    var font = GetCellFont(cell);
                    var options = new TextOptions(font);
                    var lines = text.Split(new[] { "\r\n", "\n" }, StringSplitOptions.None);
                    double lineHeight = lines.Max(line => TextMeasurer.MeasureSize(line, options).Height * pointsPerInch / options.Dpi);
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
            });
        }

        /// <summary>
        /// Freezes panes on the worksheet.
        /// </summary>
        /// <param name="topRows">Number of rows at the top to freeze.</param>
        /// <param name="leftCols">Number of columns on the left to freeze.</param>
        public void Freeze(int topRows = 0, int leftCols = 0) {
            WriteLock(() => {
                Worksheet worksheet = _worksheetPart.Worksheet;
                SheetViews sheetViews = worksheet.GetFirstChild<SheetViews>();

                if (topRows == 0 && leftCols == 0) {
                    if (sheetViews != null) {
                        worksheet.RemoveChild(sheetViews);
                    }
                    worksheet.Save();
                    return;
                }

                if (sheetViews == null) {
                    sheetViews = new SheetViews();
                    OpenXmlElement? insertBefore = worksheet.Elements<SheetFormatProperties>().Cast<OpenXmlElement>().FirstOrDefault()
                        ?? worksheet.Elements<Columns>().Cast<OpenXmlElement>().FirstOrDefault()
                        ?? worksheet.Elements<SheetData>().Cast<OpenXmlElement>().FirstOrDefault();
                    if (insertBefore != null) {
                        worksheet.InsertBefore(sheetViews, insertBefore);
                    } else {
                        worksheet.AppendChild(sheetViews);
                    }
                }

                SheetView sheetView = sheetViews.GetFirstChild<SheetView>();
                if (sheetView == null) {
                    sheetView = new SheetView { WorkbookViewId = 0U };
                    sheetViews.Append(sheetView);
                }

                sheetView.RemoveAllChildren<Pane>();
                sheetView.RemoveAllChildren<Selection>();

                Pane pane = new Pane { State = PaneStateValues.Frozen };
                if (topRows > 0) {
                    pane.HorizontalSplit = topRows;
                }
                if (leftCols > 0) {
                    pane.VerticalSplit = leftCols;
                }

                pane.TopLeftCell = GetColumnName(leftCols + 1) + (topRows + 1).ToString(CultureInfo.InvariantCulture);

                if (topRows > 0 && leftCols > 0) {
                    pane.ActivePane = PaneValues.BottomRight;
                    sheetView.Append(pane);
                    sheetView.Append(new Selection {
                        Pane = PaneValues.TopRight,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                    sheetView.Append(new Selection {
                        Pane = PaneValues.BottomLeft,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                    sheetView.Append(new Selection {
                        Pane = PaneValues.BottomRight,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                } else if (topRows > 0) {
                    pane.ActivePane = PaneValues.BottomLeft;
                    sheetView.Append(pane);
                    sheetView.Append(new Selection {
                        Pane = PaneValues.BottomLeft,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                } else {
                    pane.ActivePane = PaneValues.TopRight;
                    sheetView.Append(pane);
                    sheetView.Append(new Selection {
                        Pane = PaneValues.TopRight,
                        ActiveCell = pane.TopLeftCell,
                        SequenceOfReferences = new ListValue<StringValue> { InnerText = pane.TopLeftCell }
                    });
                }

                sheetView.Append(new Selection {
                    ActiveCell = "A1",
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
                });

                worksheet.Save();
            });
        }


        public void AddAutoFilter(string range, Dictionary<uint, IEnumerable<string>> filterCriteria = null) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            WriteLock(() => {
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
            });
        }

        /// <summary>
        /// Adds an Excel table to the worksheet over the specified range.
        /// </summary>
        /// <param name="range">Cell range (e.g. "A1:B3") defining the table area.</param>
        /// <param name="hasHeader">Indicates whether the first row is a header row.</param>
        /// <param name="name">Name of the table. If empty, a default name is used.</param>
        /// <param name="style">Table style to apply.</param>
        /// <remarks>
        /// All cells within <paramref name="range"/> must exist. Missing cells are automatically created with empty values.
        /// </remarks>
        /// <exception cref="ArgumentNullException">Thrown when <paramref name="range"/> is null or empty.</exception>
        /// <exception cref="ArgumentException">Thrown when <paramref name="range"/> is not in a valid format.</exception>
        /// <exception cref="InvalidOperationException">Thrown when the specified range overlaps with an existing table.</exception>
        public void AddTable(string range, bool hasHeader, string name, TableStyle style) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            WriteLock(() => {
                var cells = range.Split(':');
                if (cells.Length != 2) {
                    throw new ArgumentException("Invalid range format", nameof(range));
                }

                string startRef = cells[0];
                string endRef = cells[1];

                int startColumnIndex = GetColumnIndex(startRef);
                int endColumnIndex = GetColumnIndex(endRef);
                int startRowIndex = GetRowIndex(startRef);
                int endRowIndex = GetRowIndex(endRef);

                uint columnsCount = (uint)(endColumnIndex - startColumnIndex + 1);

                foreach (var existingPart in _worksheetPart.TableDefinitionParts) {
                    var existingRange = existingPart.Table?.Reference?.Value;
                    if (string.IsNullOrEmpty(existingRange)) continue;
                    var existingCells = existingRange.Split(':');
                    if (existingCells.Length != 2) continue;
                    string existingStartRef = existingCells[0];
                    string existingEndRef = existingCells[1];

                    int existingStartColumn = GetColumnIndex(existingStartRef);
                    int existingEndColumn = GetColumnIndex(existingEndRef);
                    int existingStartRow = GetRowIndex(existingStartRef);
                    int existingEndRow = GetRowIndex(existingEndRef);

                    bool overlaps = startColumnIndex <= existingEndColumn &&
                                    endColumnIndex >= existingStartColumn &&
                                    startRowIndex <= existingEndRow &&
                                    endRowIndex >= existingStartRow;
                    if (overlaps) {
                        throw new InvalidOperationException("The specified range overlaps with an existing table.");
                    }
                }

                for (int row = startRowIndex; row <= endRowIndex; row++) {
                    for (int column = startColumnIndex; column <= endColumnIndex; column++) {
                        var cell = GetCell(row, column);
                        if (cell.CellValue == null) {
                            cell.CellValue = new CellValue(string.Empty);
                            cell.DataType = CellValues.String;
                        }
                    }
                }

                uint tableId = (uint)(_worksheetPart.TableDefinitionParts.Count() + 1);
                var tableDefinitionPart = _worksheetPart.AddNewPart<TableDefinitionPart>();

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
                tableDefinitionPart.Table.Save();

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
            });
        }

        public void AddConditionalRule(string range, ConditionalFormattingOperatorValues @operator, string formula1, string? formula2 = null) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            WriteLock(() => {
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
            });
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

            WriteLock(() => {
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
            });
        }

        public void AddConditionalDataBar(string range, SixLaborsColor color) {
            AddConditionalDataBar(range, ConvertColor(color));
        }

        public void AddConditionalDataBar(string range, string color) {
            if (string.IsNullOrEmpty(range)) {
                throw new ArgumentNullException(nameof(range));
            }

            WriteLock(() => {
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
            });
        }

        public void InsertObjects<T>(IEnumerable<T> items, bool includeHeaders = true, int startRow = 1) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }

            var list = items.Cast<object?>().ToList();
            if (list.Count == 0) {
                return;
            }

            var flattenedItems = new List<Dictionary<string, object?>>();
            List<string> headers = new List<string>();

            foreach (var item in list) {
                var dict = new Dictionary<string, object?>();
                FlattenObject(item, null, dict);
                flattenedItems.Add(dict);
                foreach (var key in dict.Keys) {
                    if (!headers.Contains(key)) {
                        headers.Add(key);
                    }
                }
            }

            List<(int Row, int Column, object Value)> cells = new List<(int Row, int Column, object Value)>();
            int row = startRow;
            if (includeHeaders) {
                for (int c = 0; c < headers.Count; c++) {
                    cells.Add((row, c + 1, headers[c]));
                }
                row++;
            }

            foreach (var dict in flattenedItems) {
                for (int c = 0; c < headers.Count; c++) {
                    object value = dict.ContainsKey(headers[c]) ? dict[headers[c]] ?? string.Empty : string.Empty;
                    cells.Add((row, c + 1, value));
                }
                row++;
            }

            const int parallelThreshold = 1000;
            if (cells.Count > parallelThreshold) {
                CellValuesParallel(cells);
            } else {
                foreach (var cell in cells) {
                    CellValue(cell.Row, cell.Column, cell.Value);
                }
            }
        }

        private static void FlattenObject(object? value, string? prefix, IDictionary<string, object?> result) {
            if (value == null) {
                if (!string.IsNullOrEmpty(prefix)) {
                    result[prefix] = null;
                }
                return;
            }

            if (value is IDictionary dictionary) {
                foreach (DictionaryEntry entry in dictionary) {
                    string key = entry.Key?.ToString() ?? string.Empty;
                    string childPrefix = string.IsNullOrEmpty(prefix) ? key : prefix + "." + key;
                    FlattenObject(entry.Value, childPrefix, result);
                }
                return;
            }

            if (value is IEnumerable enumerable && value is not string) {
                var values = new List<string>();
                foreach (var item in enumerable) {
                    values.Add(item?.ToString() ?? string.Empty);
                }
                if (!string.IsNullOrEmpty(prefix)) {
                    result[prefix] = string.Join(", ", values);
                }
                return;
            }

            Type type = value.GetType();
            if (type.IsPrimitive || value is string || value is decimal || value is DateTime || value is DateTimeOffset || value is Guid) {
                if (!string.IsNullOrEmpty(prefix)) {
                    result[prefix] = value;
                }
                return;
            }

            var props = type.GetProperties().Where(p => p.CanRead);
            bool hasAny = false;
            foreach (var prop in props) {
                hasAny = true;
                string childPrefix = string.IsNullOrEmpty(prefix) ? prop.Name : prefix + "." + prop.Name;
                FlattenObject(prop.GetValue(value, null), childPrefix, result);
            }

            if (!hasAny && !string.IsNullOrEmpty(prefix)) {
                result[prefix] = value.ToString();
            }
        }

        private readonly struct CellUpdate {
            public readonly int Row;
            public readonly int Column;
            public readonly string Text;
            public readonly CellValues DataType;
            public readonly bool IsSharedString;

            public CellUpdate(int row, int column, string text, CellValues dataType, bool isSharedString) {
                Row = row;
                Column = column;
                Text = text;
                DataType = dataType;
                IsSharedString = isSharedString;
            }
        }

        /// <summary>
        /// Sets multiple cell values in parallel without mutating the DOM concurrently.
        /// Cell data is prepared in parallel, then applied sequentially under write lock
        /// to prevent XML corruption and ensure thread safety.
        /// </summary>
        /// <param name="cells">Collection of cell coordinates and values.</param>
        public void CellValuesParallel(IEnumerable<(int Row, int Column, object Value)> cells) {
            if (cells == null) {
                throw new ArgumentNullException(nameof(cells));
            }

            var cellList = cells as IList<(int Row, int Column, object Value)> ?? cells.ToList();
            int cellCount = cellList.Count;
            bool monitor = cellCount > 5000;
            Stopwatch? prepWatch = null;
            Stopwatch? applyWatch = null;
            if (monitor) {
                prepWatch = Stopwatch.StartNew();
            }

            var bag = new ConcurrentBag<CellUpdate>();

            Parallel.ForEach(cellList, cell => {
                bag.Add(PrepareCellUpdate(cell.Row, cell.Column, cell.Value));
            });

            if (monitor && prepWatch != null) {
                prepWatch.Stop();
                applyWatch = Stopwatch.StartNew();
            }

            WriteLock(() => {
                _skipWriteLock.Value = true;
                try {
                    foreach (var update in bag) {
                        ApplyCellUpdate(update);
                    }
                    ValidateWorksheetXml();
                } finally {
                    _skipWriteLock.Value = false;
                }
            });

            if (monitor && applyWatch != null && prepWatch != null) {
                applyWatch.Stop();
                Debug.WriteLine($"CellValuesParallel: prepared {cellCount} cells in {prepWatch.ElapsedMilliseconds} ms, applied in {applyWatch.ElapsedMilliseconds} ms.");
            }
        }

        private CellUpdate PrepareCellUpdate(int row, int column, object value) {
            switch (value) {
                case string s:
                    return new CellUpdate(row, column, s, CellValues.SharedString, true);
                case double d:
                    return new CellUpdate(row, column, d.ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case float f:
                    return new CellUpdate(row, column, Convert.ToDouble(f).ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case decimal dec:
                    return new CellUpdate(row, column, dec.ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case int i:
                    return new CellUpdate(row, column, ((double)i).ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case long l:
                    return new CellUpdate(row, column, ((double)l).ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case DateTime dt:
                    return new CellUpdate(row, column, dt.ToOADate().ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case DateTimeOffset dto:
                    return new CellUpdate(row, column, dto.UtcDateTime.ToOADate().ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case TimeSpan ts:
                    return new CellUpdate(row, column, ts.TotalDays.ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case bool b:
                    return new CellUpdate(row, column, b ? "1" : "0", CellValues.Boolean, false);
                case uint ui:
                    return new CellUpdate(row, column, ((double)ui).ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case ulong ul:
                    return new CellUpdate(row, column, ((double)ul).ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case ushort us:
                    return new CellUpdate(row, column, ((double)us).ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case byte by:
                    return new CellUpdate(row, column, ((double)by).ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case sbyte sb:
                    return new CellUpdate(row, column, ((double)sb).ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                case short sh:
                    return new CellUpdate(row, column, ((double)sh).ToString(CultureInfo.InvariantCulture), CellValues.Number, false);
                default:
                    return new CellUpdate(row, column, value?.ToString() ?? string.Empty, CellValues.SharedString, true);
            }
        }

        private void ApplyCellUpdate(CellUpdate update) {
            Cell cell = GetCell(update.Row, update.Column);
            if (update.IsSharedString) {
                int sharedStringIndex = _excelDocument.GetSharedStringIndex(update.Text);
                cell.CellValue = new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.SharedString;
            } else {
                cell.CellValue = new CellValue(update.Text);
                cell.DataType = update.DataType;
            }
        }

        private void ValidateWorksheetXml() {
            try {
                using StringReader sr = new StringReader(_worksheetPart.Worksheet.OuterXml);
                using XmlReader reader = XmlReader.Create(sr);
                while (reader.Read()) { }
            } catch (XmlException ex) {
                throw new InvalidOperationException($"Worksheet XML is not well-formed after parallel write operation. {ex.Message}", ex);
            }
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, string value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                int sharedStringIndex = _excelDocument.GetSharedStringIndex(value);
                cell.CellValue = new CellValue(sharedStringIndex.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.SharedString;
            });
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, double value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, decimal value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, DateTime value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.ToOADate().ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, DateTimeOffset value) {
            CellValue(row, column, value.UtcDateTime);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, TimeSpan value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value.TotalDays.ToString(CultureInfo.InvariantCulture));
                cell.DataType = CellValues.Number;
            });
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, uint value) {
            CellValue(row, column, (double)value);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, ulong value) {
            CellValue(row, column, (double)value);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, ushort value) {
            CellValue(row, column, (double)value);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, byte value) {
            CellValue(row, column, (double)value);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, sbyte value) {
            CellValue(row, column, (double)value);
        }

        /// <inheritdoc cref="CellValue(int,int,object)" />
        public void CellValue(int row, int column, bool value) {
            WriteLockConditional(() => {
                Cell cell = GetCell(row, column);
                cell.CellValue = new CellValue(value ? "1" : "0");
                cell.DataType = CellValues.Boolean;
            });
        }

        /// <summary>
        /// Sets a formula in the specified cell.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="formula">The formula expression.</param>
        public void CellFormula(int row, int column, string formula) {
            WriteLock(() => {
                Cell cell = GetCell(row, column);
                cell.CellFormula = new CellFormula(formula);
            });
        }

        /// <summary>
        /// Sets the value, formula, and number format of a cell in a single operation.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">Optional value to assign.</param>
        /// <param name="formula">Optional formula to apply.</param>
        /// <param name="numberFormat">Optional number format code.</param>
        public void Cell(int row, int column, object? value = null, string? formula = null, string? numberFormat = null) {
            if (value != null) {
                CellValue(row, column, value);
            }
            if (!string.IsNullOrEmpty(formula)) {
                CellFormula(row, column, formula);
            }
            if (!string.IsNullOrEmpty(numberFormat)) {
                FormatCell(row, column, numberFormat);
            }
        }

        /// <summary>
        /// Applies a number format to the specified cell.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="numberFormat">The number format code to apply.</param>
        public void FormatCell(int row, int column, string numberFormat) {
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

        /// <summary>
        /// Sets the value of a cell.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">The value to assign.</param>
        /// <summary>
        /// Sets the specified value into a cell, inferring the data type.
        /// </summary>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">The value to assign.</param>
        public void CellValue(int row, int column, object value) {
            WriteLockConditional(() => {
                switch (value) {
                    case string s:
                        CellValue(row, column, s);
                        break;
                    case double d:
                        CellValue(row, column, d);
                        break;
                    case float f:
                        CellValue(row, column, Convert.ToDouble(f));
                        break;
                    case decimal dec:
                        CellValue(row, column, dec);
                        break;
                    case int i:
                        CellValue(row, column, (double)i);
                        break;
                    case long l:
                        CellValue(row, column, (double)l);
                        break;
                    case DateTime dt:
                        CellValue(row, column, dt);
                        break;
                    case DateTimeOffset dto:
                        CellValue(row, column, dto);
                        break;
                    case TimeSpan ts:
                        CellValue(row, column, ts);
                        break;
                    case bool b:
                        CellValue(row, column, b);
                        break;
                    case uint ui:
                        CellValue(row, column, ui);
                        break;
                    case ulong ul:
                        CellValue(row, column, ul);
                        break;
                    case ushort us:
                        CellValue(row, column, us);
                        break;
                    case byte by:
                        CellValue(row, column, by);
                        break;
                    case sbyte sb:
                        CellValue(row, column, sb);
                        break;
                    case short sh:
                        CellValue(row, column, (double)sh);
                        break;
                    default:
                        if (value != null) {
                            CellValue(row, column, value.ToString());
                        } else {
                            CellValue(row, column, string.Empty);
                        }
                        break;
                }
            });
        }

        /// <summary>
        /// Sets the value of a cell using a nullable struct.
        /// </summary>
        /// <typeparam name="T">The value type.</typeparam>
        /// <param name="row">The 1-based row index.</param>
        /// <param name="column">The 1-based column index.</param>
        /// <param name="value">The nullable value to assign.</param>
        public void CellValue<T>(int row, int column, T? value) where T : struct {
            if (value.HasValue) {
                CellValue(row, column, value.Value);
            } else {
                CellValue(row, column, string.Empty);
            }
        }

        public void Dispose() {
            _lock.Dispose();
        }
    }
}