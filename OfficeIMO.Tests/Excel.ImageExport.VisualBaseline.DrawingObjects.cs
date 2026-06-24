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
        private const string RotatedPresetDrawingObjectBaselineName = "officeimo-excel-image-rotated-preset-drawing-object";

        [Fact]
        public void RotatedPresetDrawingObjectImageExportMatchesApprovedBaselines() {
            using ExcelBaselineFixture fixture = CreateRotatedPresetDrawingObjectBaselineWorkbook();
            ExcelRange range = fixture.Sheet.Range("A1:F8");
            ExcelImageExportOptions options = CreateBaselineOptions();

            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.DrawingShapeTextRotationApproximation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("<path", svgText, StringComparison.Ordinal);
            Assert.Contains("matrix(", svgText, StringComparison.Ordinal);
            Assert.Contains("#FDBA74", svgText, StringComparison.Ordinal);
            Assert.Contains("#EA580C", svgText, StringComparison.Ordinal);
            AssertRasterBaseline(RotatedPresetDrawingObjectBaselineName + ".png", png.Bytes);
            AssertTextBaseline(RotatedPresetDrawingObjectBaselineName + ".svg", svgText);
        }

        [Fact]
        public void ApprovedRotatedPresetDrawingObjectBaselinesAreRenderableAndNonBlank() {
            string baselineDirectory = BaselineDirectory;
            string pngPath = Path.Combine(baselineDirectory, RotatedPresetDrawingObjectBaselineName + ".png");
            string svgPath = Path.Combine(baselineDirectory, RotatedPresetDrawingObjectBaselineName + ".svg");
            if (UpdateBaselines) {
                using ExcelBaselineFixture fixture = CreateRotatedPresetDrawingObjectBaselineWorkbook();
                ExcelRange range = fixture.Sheet.Range("A1:F8");
                ExcelImageExportOptions options = CreateBaselineOptions();
                AssertRasterBaseline(RotatedPresetDrawingObjectBaselineName + ".png", range.ExportImage(OfficeImageExportFormat.Png, options).Bytes);
                AssertTextBaseline(RotatedPresetDrawingObjectBaselineName + ".svg", System.Text.Encoding.UTF8.GetString(range.ExportImage(OfficeImageExportFormat.Svg, options).Bytes));
            }

            Assert.True(File.Exists(pngPath), "Missing approved rotated preset drawing-object PNG baseline: " + pngPath);
            Assert.True(File.Exists(svgPath), "Missing approved rotated preset drawing-object SVG baseline: " + svgPath);

            OfficeRasterImage image = VisualBaselineTestSupport.DecodePng(File.ReadAllBytes(pngPath), "Approved rotated preset drawing-object PNG baseline is not a supported PNG file.");
            Assert.True(image.Width >= 600, "Rotated preset drawing-object PNG baseline width is unexpectedly small.");
            Assert.True(image.Height >= 300, "Rotated preset drawing-object PNG baseline height is unexpectedly small.");
            int fillPixels = CountPixelsNear(image, OfficeColor.FromRgb(253, 186, 116));
            Assert.True(fillPixels > 1200, "Rotated preset drawing-object PNG baseline does not contain enough visible heart fill pixels.");

            string svg = File.ReadAllText(svgPath);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("<path", svg, StringComparison.Ordinal);
            Assert.Contains("matrix(", svg, StringComparison.Ordinal);
            Assert.Contains("#FDBA74", svg, StringComparison.Ordinal);
            Assert.Contains("#EA580C", svg, StringComparison.Ordinal);
        }

        private static ExcelBaselineFixture CreateRotatedPresetDrawingObjectBaselineWorkbook() {
            string filePath = Path.Combine(Path.GetTempPath(), "OfficeIMO-ExcelRotatedPresetDrawingObjectBaseline-" + Guid.NewGuid().ToString("N") + ".xlsx");
            ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("RotatedShape");

            sheet.CellValue(1, 1, "Rotated DrawingML preset");
            sheet.Range("A1:F1").Merge();
            sheet.Range("A1:F1").SetFillColor("0F172A").SetFontColor("FFFFFF").SetBold();
            sheet.CellAlign(1, 1, HorizontalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Center);
            sheet.CellValue(7, 2, "Preset geometry, rotation, fill, outline, and clipping route through OfficeIMO.Drawing.");
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
            sheet.SetRowHeight(5, 32);
            sheet.SetRowHeight(6, 28);
            sheet.SetRowHeight(7, 30);
            sheet.SetRowHeight(8, 22);
            for (int row = 1; row <= 8; row++) {
                for (int column = 1; column <= 6; column++) {
                    sheet.CellAt(row, column).SetBorder(BorderStyleValues.Thin, "CBD5E1");
                    sheet.CellVerticalAlign(row, column, VerticalAlignmentValues.Center);
                }
            }

            AddRotatedPresetDrawingObjectShape(sheet);
            return new ExcelBaselineFixture(document, sheet);
        }

        private static void AddRotatedPresetDrawingObjectShape(ExcelSheet sheet) {
            WorksheetPart worksheetPart = sheet.WorksheetPart;
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            var transform = new A.Transform2D {
                Rotation = (int)Math.Round(28D * 60000D)
            };

            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("2"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("2"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("4"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("5"),
                    new Xdr.RowOffset("0")),
                new Xdr.Shape(
                    new Xdr.NonVisualShapeProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 120U, Name = "Rotated heart" },
                        new Xdr.NonVisualShapeDrawingProperties(new A.ShapeLocks { NoGrouping = true })),
                    new Xdr.ShapeProperties(
                        transform,
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Heart },
                        new A.SolidFill(new A.RgbColorModelHex { Val = "FDBA74" }),
                        new A.Outline(
                            new A.SolidFill(new A.RgbColorModelHex { Val = "EA580C" })) {
                            Width = 19050
                        }),
                    new Xdr.TextBody(
                        new A.BodyProperties(),
                        new A.ListStyle(),
                        new A.Paragraph())),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }
    }
}
