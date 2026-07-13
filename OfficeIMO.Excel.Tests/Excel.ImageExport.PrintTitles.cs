using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelImageExportPrintTitleTests {
        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRepeatsPrintTitleColumns() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            document.SetPrintTitles(sheet, firstRow: null, lastRow: null, firstCol: 1, lastCol: 1, save: false);
            sheet.AddManualColumnPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Equal(2, results.Count);
            Assert.Equal("Report!C1:D4", results[1].Source);
            Assert.DoesNotContain(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintTitlesUnsupported);
            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.Contains(">A1<", svg);
            Assert.Contains(">C1<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedImageExportRepeatsPrintTitleCorner() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: 1, lastCol: 1, save: false);
            sheet.AddManualRowPageBreak(2, save: false);
            sheet.AddManualColumnPageBreak(2, save: false);

            var options = new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            };
            IReadOnlyList<OfficeImageExportResult> svgResults = sheet.ExportImages(OfficeImageExportFormat.Svg, options);
            IReadOnlyList<OfficeImageExportResult> pngResults = sheet.ExportImages(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult bodyOnly = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                Range = "C3:D4",
                ShowGridlines = false
            });

            Assert.Equal("Report!C3:D4", svgResults[3].Source);
            Assert.DoesNotContain(svgResults[3].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintTitlesUnsupported);
            string svg = Encoding.UTF8.GetString(svgResults[3].Bytes);
            Assert.Contains(">A1<", svg);
            Assert.Contains(">C1<", svg);
            Assert.Contains(">A3<", svg);
            Assert.Contains(">C3<", svg);

            OfficeImageInfo composedInfo = OfficeImageReader.Identify(pngResults[3].Bytes);
            OfficeImageInfo bodyInfo = OfficeImageReader.Identify(bodyOnly.Bytes);
            Assert.True(composedInfo.Width > bodyInfo.Width);
            Assert.True(composedInfo.Height > bodyInfo.Height);
        }

        private static void FillPageBreakGrid(ExcelSheet sheet) {
            for (int row = 1; row <= 4; row++) {
                for (int column = 1; column <= 4; column++) {
                    sheet.CellValue(row, column, A1.CellReference(row, column));
                }
            }
        }
    }
}
