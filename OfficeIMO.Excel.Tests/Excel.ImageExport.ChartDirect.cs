using System.IO;
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
                widthPixels: 360,
                heightPixels: 220,
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
                widthPixels: 360,
                heightPixels: 220,
                title: "Ticket status");

            OfficeImageExportResult result = chart.ToImage()
                .AsWebp()
                .WithScale(1.25)
                .Export();

            Assert.Equal(OfficeImageExportFormat.Webp, result.Format);
            Assert.True(result.EncodedLength > 0);
        }
    }
}
