using System.IO;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportTests {
        [Theory]
        [InlineData(OfficeImageExportFormat.Png)]
        [InlineData(OfficeImageExportFormat.Svg)]
        [InlineData(OfficeImageExportFormat.Jpeg)]
        [InlineData(OfficeImageExportFormat.Tiff)]
        [InlineData(OfficeImageExportFormat.Webp)]
        public void ExcelChart_ExportsDirectlyThroughSharedImageContract(OfficeImageExportFormat format) {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Summary");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Open");
            sheet.CellValue(2, 2, 12);
            sheet.CellValue(3, 1, "Closed");
            sheet.CellValue(3, 2, 30);
            ExcelChart chart = sheet.AddChartFromRange(
                "A1:B3",
                row: 1,
                column: 4,
                width: OfficeLength.Pixels(360),
                height: OfficeLength.Pixels(220),
                title: "Ticket status");
            chart.Name = "TicketStatus";

            OfficeImageExportResult result = chart.ExportImage(format);

            Assert.Equal(format, result.Format);
            Assert.Equal("TicketStatus", result.Name);
            Assert.Equal("Summary!TicketStatus", result.Source);
            Assert.True(result.Width > 0);
            Assert.True(result.Height > 0);
            Assert.True(result.EncodedLength > 0);
        }

        [Fact]
        public void ExcelChart_UsesTheSharedFluentImageBuilder() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Summary");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Open");
            sheet.CellValue(2, 2, 12);
            ExcelChart chart = sheet.AddChartFromRange(
                "A1:B2",
                row: 1,
                column: 4,
                width: OfficeLength.Pixels(360),
                height: OfficeLength.Pixels(220),
                title: "Ticket status");

            byte[] result = chart.ToImage(new ExcelImageExportOptions { Format = OfficeImageExportFormat.Webp, Scale = 1.25D });

            Assert.Equal(OfficeImageFormat.Webp, OfficeImageReader.Identify(result).Format);
            Assert.NotEmpty(result);
        }

        [Theory]
        [InlineData(OfficeImageExportFormat.Png)]
        [InlineData(OfficeImageExportFormat.Svg)]
        public void ExcelChart_FitWithinAppliesToRasterAndSvg(OfficeImageExportFormat format) {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Summary");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Open");
            sheet.CellValue(2, 2, 12);
            ExcelChart chart = sheet.AddChartFromRange(
                "A1:B2",
                row: 1,
                column: 4,
                width: OfficeLength.Pixels(360),
                height: OfficeLength.Pixels(220),
                title: "Bounded chart");

            byte[] result = chart.ToImage(new ExcelImageExportOptions {
                Format = format, Scale = 2D, MaximumOutputWidth = 180, MaximumOutputHeight = 100
            });

            OfficeImageInfo info = OfficeImageReader.Identify(result);
            Assert.True(info.Width <= 180);
            Assert.True(info.Height <= 100);
        }

        [Theory]
        [InlineData(OfficeImageExportFormat.Png)]
        [InlineData(OfficeImageExportFormat.Svg)]
        public void ExcelChart_PreservesSmallAnchoredDimensions(OfficeImageExportFormat format) {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Summary");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Open");
            sheet.CellValue(2, 2, 12);
            ExcelChart chart = sheet.AddChartFromRange(
                "A1:B2",
                row: 1,
                column: 4,
                width: OfficeLength.Pixels(160),
                height: OfficeLength.Pixels(90),
                title: "Small chart");

            OfficeImageExportResult result = chart.ExportImage(format);

            Assert.Equal(160, result.Width);
            Assert.Equal(90, result.Height);
        }

        [Fact]
        public void ExcelChart_SvgUsesRequestedExportBackgroundWhenChartHasNoFill() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Summary");
            sheet.CellValue(1, 1, "Category");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "Open");
            sheet.CellValue(2, 2, 12);
            ExcelChart chart = sheet.AddChartFromRange(
                "A1:B2",
                row: 1,
                column: 4,
                width: OfficeLength.Pixels(320),
                height: OfficeLength.Pixels(180),
                title: "Transparent chart");
            chart.SetChartAreaStyle(noFill: true, noLine: true);

            string svg = Encoding.UTF8.GetString(chart.ExportImage(
                OfficeImageExportFormat.Svg,
                new ExcelImageExportOptions {
                    BackgroundColor = OfficeColor.FromRgb(12, 34, 56)
                }).Bytes);

            Assert.Contains("#0C2238", svg, System.StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExcelChart_DirectExportUsesAbsoluteAnchorDimensions() {
            string filePath = Path.Combine(Path.GetTempPath(), System.Guid.NewGuid() + ".xlsx");
            try {
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    ExcelSheet sheet = document.AddWorksheet("Summary");
                    sheet.CellValue(1, 1, "Category");
                    sheet.CellValue(1, 2, "Value");
                    sheet.CellValue(2, 1, "Open");
                    sheet.CellValue(2, 2, 12);
                    sheet.AddChartFromRange(
                        "A1:B2",
                        row: 1,
                        column: 4,
                        width: OfficeLength.Pixels(320),
                        height: OfficeLength.Pixels(180),
                        title: "Absolute chart");
                    document.Save();
                }

                MoveFirstChartToAbsoluteAnchor(
                    filePath,
                    xPixels: 40,
                    yPixels: 30,
                    widthPixels: 210,
                    heightPixels: 110);

                using ExcelDocument loaded = ExcelDocument.Load(filePath);
                ExcelChart chart = Assert.Single(loaded.Sheets.Single().Charts);
                OfficeImageExportResult result = chart.ExportImage(OfficeImageExportFormat.Png);

                Assert.Equal(210, result.Width);
                Assert.Equal(110, result.Height);
            } finally {
                if (File.Exists(filePath)) {
                    File.Delete(filePath);
                }
            }
        }
    }
}
