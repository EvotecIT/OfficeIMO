using System.Globalization;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using static OfficeIMO.Excel.Benchmarks.XlBenchShapeData;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Excel.Benchmarks;

internal static partial class XlBenchShapeScenarios {
    private const uint CategoryAxisId = 111111111U;
    private const uint ValueAxisId = 222222222U;

    internal static byte[] OpenXmlWrite() {
        using var output = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook)) {
            var workbookPart = document.AddWorkbookPart();
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = new Stylesheet(
                new Fonts(new Font()),
                new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
                new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())),
                new CellFormats(new CellFormat(), new CellFormat { NumberFormatId = 14, ApplyNumberFormat = true }));
            stylesPart.Stylesheet.Save();
            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            using (var writer = OpenXmlWriter.Create(worksheetPart)) {
                writer.WriteStartElement(new Worksheet());
                writer.WriteStartElement(new SheetData());
                writer.WriteStartElement(new Row());
                WriteText(writer, "Name");
                WriteText(writer, "Amount");
                WriteText(writer, "Date");
                writer.WriteEndElement();
                for (int i = 0; i < WriteRows; i++) {
                    writer.WriteStartElement(new Row());
                    WriteText(writer, WriteStrings[i]);
                    WriteNumber(writer, WriteNumbers[i], null);
                    WriteNumber(writer, WriteDates[i].ToOADate(), 1);
                    writer.WriteEndElement();
                }
                writer.WriteStartElement(new Row());
                WriteText(writer, "Total");
                writer.WriteStartElement(new Cell());
                writer.WriteElement(new CellFormula($"SUM(B2:B{WriteRows + 1})"));
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
                writer.WriteEndElement();
            }
            workbookPart.Workbook = new Workbook(new Sheets(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = "Data" }));
            workbookPart.Workbook.Save();
        }
        return output.ToArray();
    }

    private static void WriteText(OpenXmlWriter writer, string text) {
        writer.WriteStartElement(new Cell { DataType = CellValues.InlineString });
        writer.WriteStartElement(new InlineString());
        writer.WriteElement(new Text(text));
        writer.WriteEndElement();
        writer.WriteEndElement();
    }

    private static void WriteNumber(OpenXmlWriter writer, double value, uint? style) {
        var cell = new Cell { DataType = CellValues.Number };
        if (style.HasValue) {
            cell.StyleIndex = style.Value;
        }
        writer.WriteStartElement(cell);
        writer.WriteElement(new CellValue(value.ToString(CultureInfo.InvariantCulture)));
        writer.WriteEndElement();
    }

    internal static long OpenXmlReadAll() {
        using var document = SpreadsheetDocument.Open(new MemoryStream(ReadWorkbook), false);
        var workbookPart = document.WorkbookPart!;
        string[] shared = SharedStrings(workbookPart);
        long checksum = 0;
        using var reader = OpenXmlReader.Create(workbookPart.WorksheetParts.First());
        while (reader.Read()) {
            if (reader.ElementType == typeof(Cell) && reader.IsStartElement) {
                var cell = (Cell)reader.LoadCurrentElement()!;
                string raw = cell.CellValue?.InnerText ?? cell.InnerText;
                if (cell.DataType?.Value == CellValues.SharedString) {
                    raw = shared[int.Parse(raw, CultureInfo.InvariantCulture)];
                }
                checksum += raw.Length;
            }
        }
        return checksum;
    }

    internal static long OpenXmlSheets() {
        using var document = SpreadsheetDocument.Open(new MemoryStream(SheetsWorkbook), false);
        var workbookPart = document.WorkbookPart!;
        string[] shared = SharedStrings(workbookPart);
        long cells = 0;
        long text = 0;
        foreach (var worksheetPart in workbookPart.WorksheetParts) {
            using var reader = OpenXmlReader.Create(worksheetPart);
            while (reader.Read()) {
                if (reader.ElementType == typeof(Cell) && reader.IsStartElement) {
                    var cell = (Cell)reader.LoadCurrentElement()!;
                    cells++;
                    if (cell.DataType?.Value == CellValues.SharedString) {
                        text += shared[int.Parse(cell.CellValue!.InnerText, CultureInfo.InvariantCulture)].Length;
                    }
                }
            }
        }
        return cells == SheetsCellCount ? text : throw new InvalidDataException($"Read {cells} cells.");
    }

    private static string[] SharedStrings(WorkbookPart workbookPart) =>
        workbookPart.SharedStringTablePart?.SharedStringTable?.Elements<SharedStringItem>().Select(static item => item.InnerText).ToArray() ?? [];

    internal static byte[] OpenXmlReport() {
        using var output = new MemoryStream();
        using (var document = SpreadsheetDocument.Create(output, SpreadsheetDocumentType.Workbook)) {
            var workbookPart = document.AddWorkbookPart();
            workbookPart.Workbook = new Workbook();
            var stylesPart = workbookPart.AddNewPart<WorkbookStylesPart>();
            stylesPart.Stylesheet = ReportStylesheet();
            stylesPart.Stylesheet.Save();

            var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
            var worksheet = new Worksheet(
                new Columns(new Column { Min = 2, Max = 2, Width = 11.71, CustomWidth = true }),
                ReportSheetData(),
                new ConditionalFormatting(
                    new ConditionalFormattingRule(new Formula("C3>C2")) { Type = ConditionalFormatValues.Expression, FormatId = 0U, Priority = 1 },
                    new ConditionalFormattingRule(new Formula("C3<C2")) { Type = ConditionalFormatValues.Expression, FormatId = 1U, Priority = 2 }) {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = ChartPriceRange }
                });
            var drawingsPart = worksheetPart.AddNewPart<DrawingsPart>();
            var chartPart = drawingsPart.AddNewPart<ChartPart>();
            chartPart.ChartSpace = ReportChart();
            drawingsPart.WorksheetDrawing = ReportDrawing(drawingsPart.GetIdOfPart(chartPart));
            worksheet.Append(new DocumentFormat.OpenXml.Spreadsheet.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            worksheetPart.Worksheet = worksheet;
            workbookPart.Workbook.AppendChild(new Sheets(new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = ReportSheet }));
            workbookPart.Workbook.Save();
        }
        return output.ToArray();
    }

    private static Stylesheet ReportStylesheet() => new(
        new NumberingFormats(new NumberingFormat { NumberFormatId = 164, FormatCode = DateFormat }, new NumberingFormat { NumberFormatId = 165, FormatCode = PriceFormat }),
        new Fonts(new Font(), new Font(new Bold())),
        new Fills(new Fill(new PatternFill { PatternType = PatternValues.None }), new Fill(new PatternFill { PatternType = PatternValues.Gray125 })),
        new Borders(new Border(new LeftBorder(), new RightBorder(), new TopBorder(), new BottomBorder(), new DiagonalBorder())),
        new CellStyleFormats(new CellFormat()),
        new CellFormats(
            new CellFormat(),
            new CellFormat { FontId = 1, ApplyFont = true },
            new CellFormat { NumberFormatId = 164, ApplyNumberFormat = true },
            new CellFormat { NumberFormatId = 165, ApplyNumberFormat = true }),
        new DifferentialFormats(ReportDxf("FFC6EFCE", "FF006100"), ReportDxf("FFFFC7CE", "FF9C0006")) { Count = 2 });

    private static DifferentialFormat ReportDxf(string fill, string font) => new(
        new Font(new Color { Rgb = font }),
        new Fill(new PatternFill { PatternType = PatternValues.Solid, BackgroundColor = new BackgroundColor { Rgb = fill } }));

    private static SheetData ReportSheetData() {
        var data = new SheetData();
        var header = new Row { RowIndex = 1 };
        header.Append(ReportText(1, "Week"));
        header.Append(ReportText(2, "Week Ending"));
        for (int s = 0; s < StockSymbols; s++) {
            header.Append(ReportText(3 + s, Symbols[s]));
        }
        data.Append(header);
        for (int w = 0; w < StockWeeks; w++) {
            int rowIndex = 2 + w;
            var row = new Row { RowIndex = (uint)rowIndex };
            row.Append(ReportNumber(rowIndex, 1, w + 1, null));
            row.Append(ReportNumber(rowIndex, 2, WeekEndings[w].ToOADate(), 2));
            for (int s = 0; s < StockSymbols; s++) {
                row.Append(ReportNumber(rowIndex, 3 + s, Prices[w][s], 3));
            }
            data.Append(row);
        }
        return data;
    }

    private static Cell ReportText(int column, string text) => new(new InlineString(new Text(text))) {
        CellReference = ColumnName(column) + "1", DataType = CellValues.InlineString, StyleIndex = 1
    };

    private static Cell ReportNumber(int row, int column, double value, uint? style) {
        var cell = new Cell(new CellValue(value.ToString("R", CultureInfo.InvariantCulture))) {
            CellReference = ColumnName(column) + row.ToString(CultureInfo.InvariantCulture), DataType = CellValues.Number
        };
        if (style.HasValue) {
            cell.StyleIndex = style.Value;
        }
        return cell;
    }

    private static Xdr.WorksheetDrawing ReportDrawing(string relationshipId) {
        int firstColumn = 2 + StockSymbols;
        return new Xdr.WorksheetDrawing(new Xdr.TwoCellAnchor(
            new Xdr.FromMarker(new Xdr.ColumnId(firstColumn.ToString(CultureInfo.InvariantCulture)), new Xdr.ColumnOffset("0"), new Xdr.RowId("0"), new Xdr.RowOffset("0")),
            new Xdr.ToMarker(new Xdr.ColumnId((firstColumn + 14).ToString(CultureInfo.InvariantCulture)), new Xdr.ColumnOffset("0"), new Xdr.RowId("30"), new Xdr.RowOffset("0")),
            new Xdr.GraphicFrame(
                new Xdr.NonVisualGraphicFrameProperties(
                    new Xdr.NonVisualDrawingProperties { Id = 2U, Name = "Weekly prices chart" },
                    new Xdr.NonVisualGraphicFrameDrawingProperties()),
                new Xdr.Transform(new A.Offset { X = 0L, Y = 0L }, new A.Extents { Cx = 0L, Cy = 0L }),
                new A.Graphic(new A.GraphicData(new C.ChartReference { Id = relationshipId }) { Uri = "http://schemas.openxmlformats.org/drawingml/2006/chart" })),
            new Xdr.ClientData()));
    }

    private static C.ChartSpace ReportChart() {
        var lineChart = new C.LineChart(new C.Grouping { Val = C.GroupingValues.Standard }, new C.VaryColors { Val = false });
        for (int s = 0; s < StockSymbols; s++) {
            string column = ColumnName(3 + s);
            lineChart.Append(new C.LineChartSeries(
                new C.Index { Val = (uint)s },
                new C.Order { Val = (uint)s },
                new C.SeriesText(new C.StringReference { Formula = new C.Formula($"'{ReportSheet}'!${column}$1") }),
                new C.CategoryAxisData(new C.NumberReference {
                    Formula = new C.Formula($"'{ReportSheet}'!$B$2:$B${1 + StockWeeks}"),
                    NumberingCache = new C.NumberingCache(new C.FormatCode(DateFormat))
                }),
                new C.Values(new C.NumberReference { Formula = new C.Formula($"'{ReportSheet}'!${column}$2:${column}${1 + StockWeeks}") }),
                new C.Smooth { Val = false }));
        }
        lineChart.Append(new C.ShowMarker { Val = false });
        lineChart.Append(new C.AxisId { Val = CategoryAxisId });
        lineChart.Append(new C.AxisId { Val = ValueAxisId });

        var categoryAxis = new C.CategoryAxis(
            new C.AxisId { Val = CategoryAxisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Bottom },
            new C.NumberingFormat { FormatCode = DateFormat, SourceLinked = false },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = ValueAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.AutoLabeled { Val = true },
            new C.LabelAlignment { Val = C.LabelAlignmentValues.Center },
            new C.LabelOffset { Val = 100 },
            new C.NoMultiLevelLabels { Val = false });
        var valueAxis = new C.ValueAxis(
            new C.AxisId { Val = ValueAxisId },
            new C.Scaling(new C.Orientation { Val = C.OrientationValues.MinMax }),
            new C.Delete { Val = false },
            new C.AxisPosition { Val = C.AxisPositionValues.Left },
            new C.MajorGridlines(),
            new C.NumberingFormat { FormatCode = PriceFormat, SourceLinked = false },
            new C.TickLabelPosition { Val = C.TickLabelPositionValues.NextTo },
            new C.CrossingAxis { Val = CategoryAxisId },
            new C.Crosses { Val = C.CrossesValues.AutoZero },
            new C.CrossBetween { Val = C.CrossBetweenValues.Between });
        var chart = new C.Chart(
            new C.Title(new C.ChartText(new C.RichText(new A.BodyProperties(), new A.ListStyle(), new A.Paragraph(new A.Run(new A.Text(ReportTitle)))))),
            new C.AutoTitleDeleted { Val = false },
            new C.PlotArea(new C.Layout(), lineChart, categoryAxis, valueAxis),
            new C.Legend(new C.LegendPosition { Val = C.LegendPositionValues.Right }),
            new C.PlotVisibleOnly { Val = true });
        return new C.ChartSpace(new C.EditingLanguage { Val = "en-US" }, chart);
    }

    internal static double OpenXmlEdit(Stream? output) {
        using var package = new MemoryStream();
        package.Write(EditWorkbook, 0, EditWorkbook.Length);
        double total;
        using (var document = SpreadsheetDocument.Open(package, isEditable: true)) {
            var workbookPart = document.WorkbookPart!;
            var worksheet = workbookPart.WorksheetParts.First().Worksheet!;
            total = RewriteDeletedRows(worksheet.GetFirstChild<SheetData>()!);
            worksheet.GetFirstChild<SheetDimension>()?.SetAttribute(new OpenXmlAttribute("ref", string.Empty, $"A1:{ColumnName(EditSumColumn)}{EditLastRowAfterDelete}"));
            ForceRecalculation(workbookPart);
            worksheet.Save();
        }
        CopyTo(package, output);
        return total;
    }

    private static double RewriteDeletedRows(SheetData data) {
        var deleted = new HashSet<int>(RowsToDelete);
        double lastTotal = 0;
        int nextRow = 0;
        foreach (var row in data.Elements<Row>().ToList()) {
            int oldRow = (int)row.RowIndex!.Value;
            if (deleted.Contains(oldRow)) {
                row.Remove();
                continue;
            }
            nextRow++;
            bool shifted = nextRow != oldRow;
            if (shifted) {
                row.RowIndex = (uint)nextRow;
            }
            if (oldRow == 1) {
                continue;
            }
            Cell? sumCell = null;
            double total = EditColumnAValue;
            foreach (var cell in row.Elements<Cell>()) {
                int column = ColumnIndex(cell.CellReference!.Value!);
                if (shifted) {
                    cell.CellReference = ColumnName(column) + nextRow.ToString(CultureInfo.InvariantCulture);
                }
                if (column == 1) {
                    cell.DataType = CellValues.Number;
                    cell.CellValue = new CellValue(EditColumnAValue.ToString("R", CultureInfo.InvariantCulture));
                } else if (column <= EditColumns) {
                    total += Number(cell);
                } else if (column == EditSumColumn) {
                    sumCell = cell;
                }
            }
            if (sumCell != null) {
                sumCell.CellFormula = new CellFormula(EditFormula(nextRow));
                sumCell.CellValue = new CellValue(total.ToString("R", CultureInfo.InvariantCulture));
            }
            lastTotal = total;
        }
        return lastTotal;
    }

    internal static double OpenXmlInsert(Stream? output) {
        using var package = new MemoryStream();
        package.Write(EditWorkbook, 0, EditWorkbook.Length);
        double total;
        using (var document = SpreadsheetDocument.Open(package, isEditable: true)) {
            var workbookPart = document.WorkbookPart!;
            var worksheet = workbookPart.WorksheetParts.First().Worksheet!;
            total = RewriteInsertedRows(worksheet.GetFirstChild<SheetData>()!);
            worksheet.GetFirstChild<SheetDimension>()?.SetAttribute(new OpenXmlAttribute("ref", string.Empty, $"A1:{ColumnName(EditSumColumn + InsertCount)}{EditRows + 1}"));
            ForceRecalculation(workbookPart);
            worksheet.Save();
        }
        CopyTo(package, output);
        return total;
    }

    private static double RewriteInsertedRows(SheetData data) {
        double lastTotal = 0;
        foreach (var row in data.Elements<Row>()) {
            int rowIndex = (int)row.RowIndex!.Value;
            var cells = row.Elements<Cell>().Select(static cell => (Cell: cell, Column: ColumnIndex(cell.CellReference!.Value!))).ToList();
            Cell? sumCell = null;
            Cell? firstShifted = null;
            double total = InsertCount * InsertedValue;
            foreach (var (cell, column) in cells) {
                if (column >= InsertAt) {
                    firstShifted ??= cell;
                    cell.CellReference = ColumnName(column + InsertCount) + rowIndex.ToString(CultureInfo.InvariantCulture);
                }
                if (column == EditSumColumn) {
                    sumCell = cell;
                } else if (column <= EditColumns) {
                    total += Number(cell);
                }
            }
            if (rowIndex == 1) {
                continue;
            }
            for (int column = InsertAt; column < InsertAt + InsertCount; column++) {
                row.InsertBefore(new Cell {
                    CellReference = ColumnName(column) + rowIndex.ToString(CultureInfo.InvariantCulture),
                    DataType = CellValues.Number,
                    CellValue = new CellValue(InsertedValue.ToString("R", CultureInfo.InvariantCulture))
                }, firstShifted);
            }
            sumCell!.CellFormula = new CellFormula(InsertFormula(rowIndex));
            sumCell.CellValue = new CellValue(total.ToString("R", CultureInfo.InvariantCulture));
            lastTotal = total;
        }
        return lastTotal;
    }

    private static void ForceRecalculation(WorkbookPart workbookPart) {
        if (workbookPart.CalculationChainPart is { } chain) {
            workbookPart.DeletePart(chain);
        }
        var workbook = workbookPart.Workbook!;
        workbook.CalculationProperties ??= new CalculationProperties();
        workbook.CalculationProperties.FullCalculationOnLoad = true;
        workbook.Save();
    }

    private static void CopyTo(MemoryStream package, Stream? output) {
        if (output != null) {
            package.Position = 0;
            package.CopyTo(output);
        }
    }

    private static double Number(Cell cell) =>
        cell.CellValue is { } value && double.TryParse(value.InnerText, NumberStyles.Float, CultureInfo.InvariantCulture, out double number) ? number : 0d;

    private static int ColumnIndex(string reference) {
        int column = 0;
        foreach (char ch in reference) {
            if (!char.IsAsciiLetterUpper(ch)) {
                break;
            }
            column = column * 26 + ch - 'A' + 1;
        }
        return column;
    }

    internal static byte[] OpenXmlTemplate() {
        using var package = new MemoryStream();
        package.Write(TemplateWorkbook, 0, TemplateWorkbook.Length);
        using (var document = SpreadsheetDocument.Open(package, isEditable: true)) {
            var workbookPart = document.WorkbookPart!;
            var workbook = workbookPart.Workbook!;
            var sheet = workbook.Sheets!.Elements<Sheet>().First(static candidate => candidate.Name == "Report");
            var worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id!);
            var worksheet = worksheetPart.Worksheet!;
            var data = worksheet.GetFirstChild<SheetData>()!;
            var rows = data.Elements<Row>().ToDictionary(static row => (int)row.RowIndex!.Value);
            var styles = new Dictionary<int, uint>();
            foreach (var column in worksheet.GetFirstChild<Columns>()!.Elements<Column>()) {
                if (column.Style?.Value is { } style) {
                    for (int c = (int)column.Min!.Value; c <= (int)column.Max!.Value; c++) {
                        styles[c] = style;
                    }
                }
            }
            var names = workbook.DefinedNames!.Elements<DefinedName>().ToDictionary(static name => name.Name!.Value!, static name => name.Text);
            SetTemplateText(TemplateCell(rows, names["ReportTitle"]), TemplateTitle);
            var dateCell = TemplateCell(rows, names["ReportDate"]);
            dateCell.DataType = null;
            dateCell.CellValue = new CellValue(TemplateDate.ToOADate().ToString("R", CultureInfo.InvariantCulture));
            SetTemplateText(TemplateCell(rows, names["Region"]), TemplateRegion);

            for (int i = 0; i < Lines.Length; i++) {
                int rowIndex = TemplateFirstRow + i;
                TemplateLine line = Lines[i];
                Cell[] cells = [
                    Styled(TemplateNumber(rowIndex, 1, line.Date.ToOADate()), 1, styles),
                    Styled(TemplateString(rowIndex, 2, line.Product), 2, styles),
                    Styled(TemplateString(rowIndex, 3, line.Category), 3, styles),
                    Styled(TemplateNumber(rowIndex, 4, line.Quantity), 4, styles),
                    Styled(TemplateNumber(rowIndex, 5, line.UnitPrice), 5, styles),
                    Styled(new Cell { CellReference = "F" + rowIndex.ToString(CultureInfo.InvariantCulture), CellFormula = new CellFormula($"D{rowIndex}*E{rowIndex}") }, 6, styles)
                ];
                if (rows.TryGetValue(rowIndex, out var existing)) {
                    var first = existing.FirstChild;
                    foreach (var cell in cells) {
                        if (first == null) {
                            existing.Append(cell);
                        } else {
                            existing.InsertBefore(cell, first);
                        }
                    }
                } else {
                    data.Append(new Row(cells) { RowIndex = (uint)rowIndex });
                }
            }

            string tableRange = $"A5:F{TemplateFirstRow + TemplateRows - 1}";
            worksheet.GetFirstChild<SheetDimension>()!.Reference = $"A1:J{TemplateFirstRow + TemplateRows - 1}";
            var table = worksheetPart.TableDefinitionParts.First(static part => part.Table!.Name == "Sales").Table!;
            table.Reference = tableRange;
            if (table.AutoFilter is { } filter) {
                filter.Reference = tableRange;
            }
            workbook.CalculationProperties ??= new CalculationProperties();
            workbook.CalculationProperties.FullCalculationOnLoad = true;
        }
        return package.ToArray();
    }

    private static Cell TemplateCell(Dictionary<int, Row> rows, string reference) {
        var (row, column) = ParseCell(reference);
        string address = ColumnName(column) + row.ToString(CultureInfo.InvariantCulture);
        return rows[row].Elements<Cell>().First(cell => cell.CellReference == address);
    }

    private static void SetTemplateText(Cell cell, string text) {
        cell.CellValue = null;
        cell.DataType = CellValues.InlineString;
        cell.InlineString = new InlineString(new Text(text));
    }

    private static Cell TemplateString(int row, int column, string text) => new() {
        CellReference = ColumnName(column) + row.ToString(CultureInfo.InvariantCulture),
        DataType = CellValues.InlineString,
        InlineString = new InlineString(new Text(text))
    };

    private static Cell TemplateNumber(int row, int column, double value) => new() {
        CellReference = ColumnName(column) + row.ToString(CultureInfo.InvariantCulture),
        CellValue = new CellValue(value.ToString("R", CultureInfo.InvariantCulture))
    };

    private static Cell Styled(Cell cell, int column, Dictionary<int, uint> styles) {
        if (styles.TryGetValue(column, out uint style)) {
            cell.StyleIndex = style;
        }
        return cell;
    }

    internal static byte[] OpenXmlProperties() {
        using var output = new MemoryStream();
        output.Write(PropertiesWorkbook, 0, PropertiesWorkbook.Length);
        using (var document = SpreadsheetDocument.Open(output, isEditable: true)) {
            document.PackageProperties.Title = PropertyTitle;
            document.PackageProperties.Category = PropertyCategory;
        }
        return output.ToArray();
    }
}
