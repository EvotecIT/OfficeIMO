using System;
using System.IO;
using System.Linq;
using System.Threading;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Excel;
using Xunit;
using Cx = DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using Threaded = DocumentFormat.OpenXml.Office2019.Excel.ThreadedComments;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;

namespace OfficeIMO.Tests {
    public partial class Excel {
        [Fact]
        public void Test_MutationSnapshot_RestoresThreadedCommentsAndChartsheetChartExRoots() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            data.AddThreadedComment("B1", "Keep", "Tester");
            WorksheetThreadedCommentsPart threadedPart = Assert.Single(data.WorksheetPart.WorksheetThreadedCommentsParts);

            ExtendedChartPart chartPart = AddChartsheetExtendedFormula(document, "Chart", "Data!B1");
            using var cancellation = new CancellationTokenSource();

            Assert.Throws<OperationCanceledException>(() => data.ApplyTransactionalMutation(_ => {
                Assert.Single(threadedPart.ThreadedComments!.Elements<Threaded.ThreadedComment>()).Ref = "C1";
                Assert.Single(chartPart.ChartSpace!.Descendants<Cx.Formula>()).Text = "Data!C1";
                cancellation.Cancel();
            }, 0, new ExcelMutationPlanOptions(), cancellation.Token));

            Assert.Equal("B1", Assert.Single(threadedPart.ThreadedComments!
                .Elements<Threaded.ThreadedComment>()).Ref!.Value);
            Assert.Equal("Data!B1", Assert.Single(chartPart.ChartSpace!
                .Descendants<Cx.Formula>()).Text);
        }

        [Fact]
        public void Test_StructuralMutations_RejectAffectedThreeDimensionalReferencesOnly() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet first = document.AddWorksheet("Q1");
            ExcelSheet second = document.AddWorksheet("Q2");
            ExcelSheet third = document.AddWorksheet("Q3");
            ExcelSheet fourth = document.AddWorksheet("Q4");
            ExcelSheet outside = document.AddWorksheet("Outside");
            fourth.CellValue(1, 2, 1);
            ExtendedChartPart chartPart = AddChartsheetExtendedFormula(
                document,
                "Quarterly Chart",
                "'Q1:Q4'!B1");

            Assert.Contains("3-D reference", Assert.Throws<InvalidOperationException>(
                () => first.PlanInsertRows(1)).Message);
            Assert.Contains("3-D reference", Assert.Throws<InvalidOperationException>(
                () => second.PlanInsertColumns(1)).Message);
            Assert.Contains("3-D reference", Assert.Throws<InvalidOperationException>(
                () => third.PlanInsertCells("A1", ExcelCellShiftDirection.Right)).Message);
            Assert.Contains("3-D reference", Assert.Throws<InvalidOperationException>(
                () => fourth.PlanMoveRange("B1", "C1")).Message);

            Assert.NotNull(second.PlanCopyRange("A1", "B1"));
            Assert.NotNull(outside.PlanInsertColumns(1));
            Assert.Equal("'Q1:Q4'!B1", Assert.Single(chartPart.ChartSpace!
                .Descendants<Cx.Formula>()).Text);
        }

        private static ExtendedChartPart AddChartsheetExtendedFormula(
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
            ExtendedChartPart chartPart = drawingsPart.AddNewPart<ExtendedChartPart>();
            chartPart.ChartSpace = new Cx.ChartSpace(new Cx.Formula(formula));
            return chartPart;
        }
    }
}
