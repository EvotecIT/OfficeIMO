using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelImageExportPatternFillTests {
        [Fact]
        public void ExcelRange_ImageExportRendersPatternFillAndReportsApproximation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Patterns");
            sheet.CellValue(1, 1, "Pattern");
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 42);
            ApplyPatternFill(sheet, 1, 1, PatternValues.DarkGrid, "FFC00000", "FFFFE5E5");

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });

            ExcelVisualCell cell = Assert.Single(snapshot.Cells);
            Assert.Equal("darkGrid", cell.Style.FillPatternType);
            Assert.Equal("FFC00000", cell.Style.FillPatternForegroundColorArgb);
            Assert.Equal("FFFFE5E5", cell.Style.FillPatternBackgroundColorArgb);
            Assert.Contains("xl-fill-1-1", svg, StringComparison.Ordinal);
            Assert.Contains("stroke=\"#C00000\"", svg, StringComparison.Ordinal);
            Assert.Contains(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.FillPatternApproximation && diagnostic.Source == "Patterns!A1");
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.NotNull(image);
            Assert.True(CountRedPixels(image!) > 20);
            Assert.True(CountPaleRedPixels(image!) > 100);
        }

        [Fact]
        public void ExcelRange_ImageExportRendersGrayPatternFillsAsSharedPercentStipples() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Patterns");
            sheet.CellValue(1, 1, "Gray");
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 42);
            ApplyPatternFill(sheet, 1, 1, PatternValues.Gray125, "FF00A000", "FFFFFFFF");

            ExcelRange range = sheet.Range("A1:A1");
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });

            Assert.Contains("fill=\"#00A000\"", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("stroke=\"#00A000\"", svg, StringComparison.Ordinal);
            Assert.Contains(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.FillPatternApproximation && diagnostic.Source == "Patterns!A1");
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.NotNull(image);
            Assert.True(CountGreenPixels(image!) > 20);
        }

        [Fact]
        public void ExcelRange_ImageExportRendersSimpleLinearGradientFill() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Gradients");
            sheet.CellValue(1, 1, "Gradient");
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 42);
            ApplyGradientFill(sheet, 1, 1, "FF0000FF", "FF00FF00", 0D);

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });

            ExcelVisualCell cell = Assert.Single(snapshot.Cells);
            Assert.Equal("gradient", cell.Style.FillPatternType);
            Assert.False(cell.Style.FillGradientUnsupported);
            Assert.Equal("FF0000FF", cell.Style.FillGradientStartColorArgb);
            Assert.Equal("FF00FF00", cell.Style.FillGradientEndColorArgb);
            Assert.Equal(0D, cell.Style.FillGradientDegree);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.FillGradientUnsupported);
            Assert.Contains("xl-gradient-1-1", svg, StringComparison.Ordinal);
            Assert.Contains("stop-color=\"#0000FF\"", svg, StringComparison.Ordinal);
            Assert.Contains("stop-color=\"#00FF00\"", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.NotNull(image);
            Assert.True(CountBluePixels(image!) > 20);
            Assert.True(CountGreenPixels(image!) > 20);
        }

        [Fact]
        public void ExcelRange_ImageExportRendersMultiStopLinearGradientFill() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Gradients");
            sheet.CellValue(1, 1, "Multi");
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 42);
            ApplyGradientFill(sheet, 1, 1, "FF0000FF", "FF00FF00", "FFFF0000", 0D);

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });

            ExcelVisualCell cell = Assert.Single(snapshot.Cells);
            Assert.False(cell.Style.FillGradientUnsupported);
            Assert.Equal(3, cell.Style.FillGradientStops.Count);
            Assert.Equal(0.5D, cell.Style.FillGradientStops[1].Offset);
            Assert.Equal("FF00FF00", cell.Style.FillGradientStops[1].ColorArgb);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.FillGradientUnsupported);
            Assert.Contains("offset=\"50%\" stop-color=\"#00FF00\"", svg, StringComparison.Ordinal);
            Assert.Contains("stop-color=\"#0000FF\"", svg, StringComparison.Ordinal);
            Assert.Contains("stop-color=\"#FF0000\"", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.NotNull(image);
            Assert.True(CountBluePixels(image!) > 20);
            Assert.True(CountGreenPixels(image!) > 20);
            Assert.True(CountRedPixels(image!) > 20);
        }

        [Fact]
        public void ExcelRange_ImageExportClampsAcceptedGradientEndpointOffsets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Gradients");
            sheet.CellValue(1, 1, "Near");
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 42);
            ApplyGradientFill(sheet, 1, 1, 0.0000005D, 0.5D, 0.9999995D);

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            string svg = range.ToSvg(new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });

            ExcelVisualCell cell = Assert.Single(snapshot.Cells);
            Assert.False(cell.Style.FillGradientUnsupported);
            Assert.Equal(0D, cell.Style.FillGradientStops[0].Offset);
            Assert.Equal(1D, cell.Style.FillGradientStops[cell.Style.FillGradientStops.Count - 1].Offset);
            Assert.Contains("offset=\"0%\" stop-color=\"#0000FF\"", svg, StringComparison.Ordinal);
            Assert.Contains("offset=\"100%\" stop-color=\"#FF0000\"", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnsupportedGradientFill() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Gradients");
            sheet.CellValue(1, 1, "Gradient");
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 42);
            ApplyGradientFill(sheet, 1, 1);

            OfficeImageExportResult png = sheet.Range("A1:A1").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });

            Assert.Contains(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.FillGradientUnsupported && diagnostic.Source == "Gradients!A1");
        }

        private static int CountRedPixels(OfficeRasterImage image) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.R > 140 && pixel.G < 80 && pixel.B < 80 && pixel.A > 180) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static int CountPaleRedPixels(OfficeRasterImage image) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.R > 230 && pixel.G > 180 && pixel.B > 180 && pixel.A > 180) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static int CountBluePixels(OfficeRasterImage image) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.B > 140 && pixel.R < 100 && pixel.G < 130 && pixel.A > 180) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static int CountGreenPixels(OfficeRasterImage image) {
            int count = 0;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.G > 140 && pixel.R < 100 && pixel.B < 130 && pixel.A > 180) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static void ApplyPatternFill(ExcelSheet sheet, int row, int column, PatternValues pattern, string foregroundArgb, string backgroundArgb) {
            ApplyFill(sheet, row, column, new Fill(new PatternFill {
                PatternType = pattern,
                ForegroundColor = new ForegroundColor { Rgb = foregroundArgb },
                BackgroundColor = new BackgroundColor { Rgb = backgroundArgb }
            }));
        }

        private static void ApplyGradientFill(ExcelSheet sheet, int row, int column) {
            ApplyFill(sheet, row, column, new Fill(new GradientFill {
                Degree = 0D
            }));
        }

        private static void ApplyGradientFill(ExcelSheet sheet, int row, int column, string startArgb, string endArgb, double degree) {
            ApplyFill(sheet, row, column, new Fill(new GradientFill(
                new GradientStop(new Color { Rgb = startArgb }) { Position = 0D },
                new GradientStop(new Color { Rgb = endArgb }) { Position = 1D }) {
                Degree = degree
            }));
        }

        private static void ApplyGradientFill(ExcelSheet sheet, int row, int column, string startArgb, string middleArgb, string endArgb, double degree) {
            ApplyFill(sheet, row, column, new Fill(new GradientFill(
                new GradientStop(new Color { Rgb = startArgb }) { Position = 0D },
                new GradientStop(new Color { Rgb = middleArgb }) { Position = 0.5D },
                new GradientStop(new Color { Rgb = endArgb }) { Position = 1D }) {
                Degree = degree
            }));
        }

        private static void ApplyGradientFill(ExcelSheet sheet, int row, int column, double startOffset, double middleOffset, double endOffset) {
            ApplyFill(sheet, row, column, new Fill(new GradientFill(
                new GradientStop(new Color { Rgb = "FF0000FF" }) { Position = startOffset },
                new GradientStop(new Color { Rgb = "FF00FF00" }) { Position = middleOffset },
                new GradientStop(new Color { Rgb = "FFFF0000" }) { Position = endOffset }) {
                Degree = 0D
            }));
        }

        private static void ApplyFill(ExcelSheet sheet, int row, int column, Fill fill) {
            WorkbookPart workbookPart = sheet.WorksheetPart.GetParentParts().OfType<WorkbookPart>().Single();
            WorkbookStylesPart stylesPart = workbookPart.WorkbookStylesPart ?? workbookPart.AddNewPart<WorkbookStylesPart>();
            Stylesheet stylesheet = stylesPart.Stylesheet ??= new Stylesheet();
            stylesheet.Fills ??= new Fills(
                new Fill(new PatternFill { PatternType = PatternValues.None }),
                new Fill(new PatternFill { PatternType = PatternValues.Gray125 }));
            stylesheet.CellFormats ??= new CellFormats(new CellFormat());

            stylesheet.Fills.Append(fill);
            uint fillId = (uint)stylesheet.Fills.Count();
            stylesheet.Fills.Count = fillId;

            Cell cell = sheet.WorksheetPart.Worksheet!.Descendants<Cell>()
                .Single(item => string.Equals(item.CellReference?.Value, A1.CellReference(row, column), StringComparison.OrdinalIgnoreCase));
            CellFormat baseFormat = stylesheet.CellFormats.Elements<CellFormat>().ElementAtOrDefault((int)(cell.StyleIndex?.Value ?? 0U)) ?? new CellFormat();
            CellFormat format = (CellFormat)baseFormat.CloneNode(true);
            format.FillId = fillId - 1U;
            format.ApplyFill = true;
            stylesheet.CellFormats.Append(format);
            uint styleIndex = (uint)stylesheet.CellFormats.Count();
            stylesheet.CellFormats.Count = styleIndex;
            cell.StyleIndex = styleIndex - 1U;
            stylesheet.Save();
        }
    }
}
