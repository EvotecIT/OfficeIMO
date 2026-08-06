using System.IO;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_TableSchema_RewritesWorksheetChartExAndChartsheetFormulaRoots() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Amount");
            sheet.CellValue(2, 1, 10);
            sheet.AddTable("A1:A2", true, "Sales", OfficeIMO.Excel.ExcelTableStyle.TableStyleMedium2);

            DrawingsPart worksheetDrawings = sheet.WorksheetPart.AddNewPart<DrawingsPart>();
            worksheetDrawings.WorksheetDrawing = new Xdr.WorksheetDrawing();
            ExtendedChartPart extendedPart = worksheetDrawings.AddNewPart<ExtendedChartPart>();
            var extendedFormula = new Cx.Formula("SUM(Sales[Amount])");
            extendedPart.ChartSpace = new Cx.ChartSpace(extendedFormula);

            C.Formula chartsheetFormula = AddChartsheetClassicFormula(
                document,
                "Chart",
                "SUM(Sales[Amount])");

            Assert.Equal("Orders", sheet.RenameTable("Sales", "Orders"));
            sheet.SetTableSchema("Orders", new[] { "Net" });

            Assert.Equal("SUM(Orders[Net])", extendedFormula.Text);
            Assert.Equal("SUM(Orders[Net])", chartsheetFormula.Text);
        }

        [Fact]
        public void Test_FormulaSyntaxTree_ParsesExternalQualifiedNamesBeforeStructuredSelectors() {
            const string formula = "[Book.xlsx]Sheet1!Rate+[Book.xlsx]Sheet1!TaxRate";

            ExcelFormulaSyntaxTree tree = ExcelFormulaSyntaxTree.Parse(formula);

            Assert.Equal(
                new[] { "[Book.xlsx]Sheet1!Rate", "[Book.xlsx]Sheet1!TaxRate" },
                tree.Nodes.OfType<ExcelFormulaNameSyntax>().Select(node => node.Name).ToArray());
            Assert.Empty(tree.Nodes.OfType<ExcelFormulaStructuredReferenceSyntax>());
            Assert.Equal(
                "[Book.xlsx]Sheet1!Rate2026+[Book.xlsx]Sheet1!TaxRate",
                tree.RewriteNames(name => name == "[Book.xlsx]Sheet1!Rate"
                    ? "[Book.xlsx]Sheet1!Rate2026"
                    : name));
            Assert.Equal(formula, tree.RewriteStructuredReferences((_, _) => "BROKEN"));
        }

        [Fact]
        public void Test_StructuralColumns_RemoveStaleRowSpans() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet insertSheet = document.AddWorksheet("Insert");
            insertSheet.CellValue(1, 1, "A");
            insertSheet.CellValue(1, 2, "B");
            insertSheet.CellValue(1, 3, "C");
            insertSheet.CellValue(2, 1, "Unchanged");
            Row[] insertRows = insertSheet.WorksheetPart.Worksheet!
                .GetFirstChild<SheetData>()!.Elements<Row>().ToArray();
            Row insertedRow = insertRows.Single(row => row.RowIndex!.Value == 1U);
            Row unchangedRow = insertRows.Single(row => row.RowIndex!.Value == 2U);
            insertedRow.Spans = new ListValue<StringValue> { InnerText = "1:3" };
            unchangedRow.Spans = new ListValue<StringValue> { InnerText = "1:1" };

            insertSheet.InsertColumns(2);

            Assert.Null(insertedRow.Spans);
            Assert.Equal("1:1", unchangedRow.Spans!.InnerText);
            Assert.Equal(new[] { "A1", "C1", "D1" },
                insertedRow.Elements<Cell>().Select(cell => cell.CellReference!.Value).ToArray());

            ExcelSheet deleteSheet = document.AddWorksheet("Delete");
            deleteSheet.CellValue(1, 1, "A");
            deleteSheet.CellValue(1, 2, "B");
            deleteSheet.CellValue(1, 3, "C");
            Row deletedRow = Assert.Single(deleteSheet.WorksheetPart.Worksheet!
                .GetFirstChild<SheetData>()!.Elements<Row>());
            deletedRow.Spans = new ListValue<StringValue> { InnerText = "1:3" };

            deleteSheet.DeleteColumns(1);

            Assert.Null(deletedRow.Spans);
            Assert.Equal(new[] { "A1", "B1" },
                deletedRow.Elements<Cell>().Select(cell => cell.CellReference!.Value).ToArray());
        }

        private static C.Formula AddChartsheetClassicFormula(
            ExcelDocument document,
            string sheetName,
            string formula) {
            ChartsheetPart chartsheetPart = document.WorkbookPartRoot.AddNewPart<ChartsheetPart>();
            DrawingsPart drawingsPart = chartsheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            chartsheetPart.Chartsheet = new Chartsheet(
                new DocumentFormat.OpenXml.Spreadsheet.Drawing {
                    Id = chartsheetPart.GetIdOfPart(drawingsPart)
                });
            uint nextSheetId = document.WorkbookRoot.Sheets!.Elements<Sheet>()
                .Max(sheet => sheet.SheetId?.Value ?? 0U) + 1U;
            document.WorkbookRoot.Sheets.Append(new Sheet {
                Id = document.WorkbookPartRoot.GetIdOfPart(chartsheetPart),
                SheetId = nextSheetId,
                Name = sheetName
            });

            ChartPart chartPart = drawingsPart.AddNewPart<ChartPart>();
            var chartFormula = new C.Formula(formula);
            chartPart.ChartSpace = new C.ChartSpace(
                new C.Chart(
                    new C.PlotArea(
                        new C.LineChart(
                            new C.Grouping { Val = C.GroupingValues.Standard },
                            new C.LineChartSeries(
                                new C.Index { Val = 0U },
                                new C.Order { Val = 0U },
                                new C.Values(new C.NumberReference(chartFormula)))))));
            return chartFormula;
        }
    }
}
