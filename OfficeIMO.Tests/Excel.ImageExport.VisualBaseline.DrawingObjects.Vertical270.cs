using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using A = DocumentFormat.OpenXml.Drawing;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportVisualBaselineTests {
        private const string Vertical270ShapeTextBaselineName = "officeimo-excel-image-vertical270-shape-text";

        [Fact]
        public void Vertical270ShapeTextImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateVertical270ShapeTextBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:F8");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextVerticalOrientationUnsupported);
            Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
            Assert.Single(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("V270", svgText, StringComparison.Ordinal);
            Assert.Contains("transform=\"rotate(270", svgText, StringComparison.Ordinal);
            Assert.Contains("#FEF3C7", svgText, StringComparison.Ordinal);
            Assert.Contains("#D97706", svgText, StringComparison.Ordinal);
            AssertDiagnosticsBaseline(Vertical270ShapeTextBaselineName + ".diagnostics.txt", png.Diagnostics);
            AssertRasterBaseline(Vertical270ShapeTextBaselineName + ".png", png.Bytes);
            AssertTextBaseline(Vertical270ShapeTextBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedVertical270ShapeTextBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, Vertical270ShapeTextBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, Vertical270ShapeTextBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateVertical270ShapeTextBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:F8");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(Vertical270ShapeTextBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(Vertical270ShapeTextBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved vertical270 shape-text PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved vertical270 shape-text SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved vertical270 shape-text PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 600, "Vertical270 shape-text PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 300, "Vertical270 shape-text PNG baseline height is unexpectedly small.");
            int fillPixels = CountPixelsNear(image, OfficeColor.FromRgb(254, 243, 199));
            int textPixels = CountPixelsNear(image, OfficeColor.FromRgb(31, 41, 55));
            Assert.True(fillPixels > 1200, "Vertical270 shape-text PNG baseline does not contain enough visible shape fill pixels.");
            Assert.True(textPixels > 30, "Vertical270 shape-text PNG baseline does not contain enough visible rotated text pixels.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("V270", svg, StringComparison.Ordinal);
            Assert.Contains("transform=\"rotate(270", svg, StringComparison.Ordinal);
            Assert.Contains("#FEF3C7", svg, StringComparison.Ordinal);
            Assert.Contains("#D97706", svg, StringComparison.Ordinal);
        }

        private static ExcelBaselineFixture CreateVertical270ShapeTextBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelVertical270ShapeTextBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Vertical270Text");

            sheet.CellValue(1, 1, "Vertical270 DrawingML text");
            sheet.Range("A1:F1").Merge();
            sheet.Range("A1:F1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(7, 2, "Vertical270 shape text routes through shared rotated text layout.");
            sheet.Range("B7:E7").Merge();
            sheet.Range("B7:E7").SetFillColor("F8FAFC").SetFontColor("334155");
            sheet.CellAlign(7, 2, HorizontalAlignmentValues.Center);

            for (int column = 1; column <= 6; column++) {
                sheet.SetColumnWidth(column, column == 1 || column == 6 ? 9 : 15);
            }

            sheet.SetRowHeight(1, 26);
            sheet.SetRowHeight(2, 42);
            sheet.SetRowHeight(3, 42);
            sheet.SetRowHeight(4, 42);
            sheet.SetRowHeight(5, 42);
            sheet.SetRowHeight(6, 22);
            sheet.SetRowHeight(7, 30);
            sheet.SetRowHeight(8, 22);
            for (int row = 1; row <= 8; row++) {
                for (int column = 1; column <= 6; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddVertical270ShapeTextObject(sheet);
            return new ExcelBaselineFixture(document, sheet);
        }

        private static void AddVertical270ShapeTextObject(ExcelSheet sheet) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("2"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("2"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("4"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("6"),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 124U, Name = "Vertical270 text box" },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(),
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.RoundRectangle },
                        new A.SolidFill(new A.RgbColorModelHex { Val = "FEF3C7" }),
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex { Val = "D97706" })) {
                            Width = 19050
                        }),
                    new Xdr.TextBody(
                        new A.BodyProperties {
                            Anchor = A.TextAnchoringTypeValues.Center,
                            Vertical = A.TextVerticalValues.Vertical270
                        },
                        new A.ListStyle(),
                        new A.Paragraph(
                            new A.ParagraphProperties { Alignment = A.TextAlignmentTypeValues.Center },
                            new A.Run(
                                new A.RunProperties { FontSize = 1400 },
                                new A.Text("V270"))))),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }
    }
}
