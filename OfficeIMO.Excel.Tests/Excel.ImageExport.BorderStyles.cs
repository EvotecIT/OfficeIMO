using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelImageExportBorderStyleTests {
        [Fact]
        public void ExcelRange_ImageExportPreservesPremiumBorderStylesInSvgAndPng() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Borders");
            for (int column = 1; column <= 5; column++) {
                sheet.CellValue(1, column, "B" + column.ToString(System.Globalization.CultureInfo.InvariantCulture));
                sheet.SetColumnWidth(column, 11);
            }

            sheet.SetRowHeight(1, 30);
            sheet.CellAt(1, 1).SetBorder(BorderStyleValues.Dashed, "C00000");
            sheet.CellAt(1, 2).SetBorder(BorderStyleValues.Dotted, "00A000");
            sheet.CellAt(1, 3).SetBorder(BorderStyleValues.MediumDashDotDot, "004C99");
            sheet.CellAt(1, 4).SetBorder(BorderStyleValues.Double, "C00000");
            sheet.CellAt(1, 5).SetDiagonalBorder(BorderStyleValues.DashDotDot, "004C99", diagonalUp: true, diagonalDown: true);

            ExcelRange range = sheet.Range("A1:E1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { Scale = 2, ShowGridlines = false });

            Assert.Contains("stroke-dasharray=\"4 2\"", svg, StringComparison.Ordinal);
            Assert.Contains("stroke-dasharray=\"1 2\" stroke-linecap=\"round\"", svg, StringComparison.Ordinal);
            Assert.Contains("stroke-dasharray=\"8 4 2 4 2 4\"", svg, StringComparison.Ordinal);
            Assert.True(CountOccurrences(svg, "#C00000") >= 6);
            Assert.True(CountOccurrences(svg, "#004C99") >= 6);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);

            ExcelVisualCell doubleBorderCell = snapshot.Cells.Single(cell => cell.Column == 4);
            int borderX = (int)Math.Round(doubleBorderCell.X * 2D);
            int sampleY = (int)Math.Round((doubleBorderCell.Y + (doubleBorderCell.Height / 2D)) * 2D);
            Assert.True(IsRed(rendered!, borderX - 3, sampleY));
            Assert.True(IsRed(rendered!, borderX + 3, sampleY));
        }

        private static int CountOccurrences(string text, string value) {
            int count = 0;
            int index = 0;
            while ((index = text.IndexOf(value, index, StringComparison.Ordinal)) >= 0) {
                count++;
                index += value.Length;
            }

            return count;
        }

        private static bool IsRed(OfficeRasterImage image, int x, int y) {
            OfficeColor pixel = image.GetPixel(x, y);
            return pixel.R > 160 && pixel.G < 80 && pixel.B < 80 && pixel.A > 180;
        }
    }
}
