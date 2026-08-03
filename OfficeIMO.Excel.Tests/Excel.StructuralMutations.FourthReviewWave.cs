using System;
using System.IO;
using System.Linq;
using System.Xml.Linq;
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
        public void Test_StructuralColumns_RewritesChartsheetAndExtendedChartReferencesAndCaches() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet data = document.AddWorksheet("Data");
            data.CellValue(1, 2, 1);
            data.CellValue(1, 3, 2);

            ChartsheetPart chartsheetPart = document.WorkbookPartRoot.AddNewPart<ChartsheetPart>();
            DrawingsPart drawingsPart = chartsheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing = new Xdr.WorksheetDrawing();
            chartsheetPart.Chartsheet = new Chartsheet(
                new DocumentFormat.OpenXml.Spreadsheet.Drawing {
                    Id = chartsheetPart.GetIdOfPart(drawingsPart)
                });
            Sheets sheets = document.WorkbookRoot.Sheets!;
            uint nextSheetId = sheets.Elements<Sheet>().Max(sheet => sheet.SheetId?.Value ?? 0U) + 1U;
            sheets.Append(new Sheet {
                Id = document.WorkbookPartRoot.GetIdOfPart(chartsheetPart),
                SheetId = nextSheetId,
                Name = "Chart"
            });

            ChartPart classicPart = drawingsPart.AddNewPart<ChartPart>();
            var classicFormula = new C.Formula("Data!$B$1:$C$1");
            var classicCache = new C.NumberingCache(
                new C.PointCount { Val = 2U },
                new C.NumericPoint(new C.NumericValue("1")) { Index = 0U },
                new C.NumericPoint(new C.NumericValue("2")) { Index = 1U });
            var classicReference = new C.NumberReference(classicFormula, classicCache);
            classicPart.ChartSpace = new C.ChartSpace(
                new C.Chart(
                    new C.PlotArea(
                        new C.LineChart(
                            new C.Grouping { Val = C.GroupingValues.Standard },
                            new C.LineChartSeries(
                                new C.Index { Val = 0U },
                                new C.Order { Val = 0U },
                                new C.Values(classicReference))))));

            ExtendedChartPart extendedPart = drawingsPart.AddNewPart<ExtendedChartPart>();
            var extendedFormula = new Cx.Formula("Data!B1:C1");
            var level = new Cx.NumericLevel(
                new Cx.NumericValue("1") { Idx = 0U },
                new Cx.NumericValue("2") { Idx = 1U }) { PtCount = 2U };
            var dimension = new Cx.NumericDimension(extendedFormula, level);
            extendedPart.ChartSpace = new Cx.ChartSpace(dimension);

            data.InsertColumns(2);

            Assert.Equal("Data!$C$1:$D$1", classicFormula.Text);
            Assert.Empty(classicReference.Elements<C.NumberingCache>());
            Assert.Equal("Data!C1:D1", extendedFormula.Text);
            Assert.Empty(dimension.Elements<Cx.NumericLevel>());
        }

        [Fact]
        public void Test_StructuralColumns_RemapsConnectionParametersAndRejectsTheirDeletion() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            Parameter parameter = AttachCellBackedConnection(document, sheet, "A5");

            sheet.InsertColumns(1);
            Assert.Equal("B5", parameter.Cell!.Value);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                () => sheet.DeleteColumns(2));
            Assert.Contains("connection parameter", exception.Message, StringComparison.OrdinalIgnoreCase);
            Assert.Equal("B5", parameter.Cell!.Value);
        }

        [Fact]
        public void Test_StructuralColumns_RemapsCommentsAndVmlAnchors() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.SetComment(1, 2, "Shift me", author: "Tester");
            VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
            int[] before = ReadCommentVmlAnchor(vmlPart);

            sheet.InsertColumns(2);

            Assert.True(sheet.HasComment(1, 3));
            Assert.False(sheet.HasComment(1, 2));
            int[] after = ReadCommentVmlAnchor(vmlPart);
            Assert.Equal(before[0] + 1, after[0]);
            Assert.Equal(before[4] + 1, after[4]);

            sheet.DeleteColumns(3);
            Assert.False(sheet.HasComment(1, 3));
            Assert.Null(sheet.WorksheetPart.WorksheetCommentsPart);
            Assert.Empty(sheet.WorksheetPart.VmlDrawingParts);
        }

        [Fact]
        public void Test_StructuralColumns_PreflightsUnsupportedAndOverflowingVmlAnchors() {
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.CellValue(1, 1, "Keep");
                sheet.WorksheetPart.Worksheet.Append(new Controls());

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                    () => sheet.InsertColumns(1));

                Assert.Contains("form controls", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.Equal("Keep", sheet.CellAt(1, 1).GetValue<string>());
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = document.AddWorksheet("Data");
                sheet.SetComment(1, 2, "Keep", author: "Tester");
                VmlDrawingPart vmlPart = Assert.Single(sheet.WorksheetPart.VmlDrawingParts);
                SetCommentVmlAnchor(vmlPart, "16383, 15, 0, 2, 16384, 15, 3, 4");

                InvalidOperationException exception = Assert.Throws<InvalidOperationException>(
                    () => sheet.InsertColumns(A1.MaxColumns));

                Assert.Contains("comment note anchor", exception.Message, StringComparison.OrdinalIgnoreCase);
                Assert.True(sheet.HasComment(1, 2));
            }
        }

        [Fact]
        public void Test_StructuralMutations_MaterializeSharedFormulasBeforeDeletingTheirMaster() {
            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = CreateHorizontalSharedFormulaSheet(document);

                sheet.DeleteColumns(2);

                AssertMaterializedSurvivingFormulas(sheet);
            }

            using (var document = ExcelDocument.Create(new MemoryStream())) {
                ExcelSheet sheet = CreateHorizontalSharedFormulaSheet(document);

                sheet.DeleteCells("B1", ExcelCellShiftDirection.Left);

                AssertMaterializedSurvivingFormulas(sheet);
            }
        }

        [Fact]
        public void Test_RangeMutations_RejectR1C1WorkbookMode() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            document.WorkbookRoot.Append(new CalculationProperties {
                ReferenceMode = ReferenceModeValues.R1C1
            });

            InvalidOperationException shift = Assert.Throws<InvalidOperationException>(
                () => sheet.PlanInsertCells("A1", ExcelCellShiftDirection.Right));
            InvalidOperationException copy = Assert.Throws<InvalidOperationException>(
                () => sheet.PlanCopyRange("A1", "B1"));

            Assert.Contains("R1C1", shift.Message);
            Assert.Contains("R1C1", copy.Message);
        }

        [Fact]
        public void Test_RangeMutationPlanning_EnforcesScanBudgets() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 2);

            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                sheet.PlanInsertCells(
                    "A1",
                    ExcelCellShiftDirection.Down,
                    new ExcelMutationPlanOptions {
                        MaximumScannedElements = 1,
                        MaximumScannedCharacters = 1_000_000
                    }));

            Assert.Contains("MaximumScannedElements", exception.Message);
            Assert.Equal(1, sheet.CellAt(1, 1).GetValue<int>());
            Assert.Equal(2, sheet.CellAt(2, 1).GetValue<int>());
        }

        [Fact]
        public void Test_CopyRange_EmitsRefWhenRelativeFormulaCrossesWorksheetBoundary() {
            using var document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellFormula(1, 2, "A1");

            sheet.CopyRange("B1", "A1");

            Assert.Equal("#REF!", sheet.GetFormulaText(1, 1));
        }

        private static ExcelSheet CreateHorizontalSharedFormulaSheet(ExcelDocument document) {
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, 1);
            Row row = Assert.Single(sheet.WorksheetPart.Worksheet.GetFirstChild<SheetData>()!.Elements<Row>());
            row.Append(
                new Cell {
                    CellReference = "B1",
                    CellFormula = new CellFormula("$A1*2") {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 91U,
                        Reference = "B1:D1"
                    }
                },
                new Cell {
                    CellReference = "C1",
                    CellFormula = new CellFormula {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 91U
                    }
                },
                new Cell {
                    CellReference = "D1",
                    CellFormula = new CellFormula {
                        FormulaType = CellFormulaValues.Shared,
                        SharedIndex = 91U
                    }
                });
            return sheet;
        }

        private static void AssertMaterializedSurvivingFormulas(ExcelSheet sheet) {
            Cell[] cells = sheet.WorksheetPart.Worksheet.Descendants<Cell>()
                .Where(cell => cell.CellFormula != null)
                .OrderBy(cell => cell.CellReference?.Value)
                .ToArray();
            Assert.Equal(new[] { "B1", "C1" }, cells.Select(cell => cell.CellReference!.Value).ToArray());
            Assert.Equal(new[] { "$A1*2", "$A1*2" }, cells.Select(cell => cell.CellFormula!.Text).ToArray());
            Assert.All(cells, cell => {
                Assert.Null(cell.CellFormula!.FormulaType);
                Assert.Null(cell.CellFormula.SharedIndex);
                Assert.True(cell.CellFormula.CalculateCell!.Value);
            });
        }

        private static int[] ReadCommentVmlAnchor(VmlDrawingPart part) {
            using Stream stream = part.GetStream();
            XDocument document = XDocument.Load(stream);
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            return Assert.Single(document.Descendants(x + "Anchor"))
                .Value.Split(',')
                .Select(value => int.Parse(value.Trim(), System.Globalization.CultureInfo.InvariantCulture))
                .ToArray();
        }

        private static void SetCommentVmlAnchor(VmlDrawingPart part, string value) {
            XDocument document;
            using (Stream stream = part.GetStream()) document = XDocument.Load(stream);
            XNamespace x = "urn:schemas-microsoft-com:office:excel";
            Assert.Single(document.Descendants(x + "Anchor")).Value = value;
            using Stream output = part.GetStream(FileMode.Create, FileAccess.Write);
            document.Save(output, SaveOptions.DisableFormatting);
        }
    }
}
