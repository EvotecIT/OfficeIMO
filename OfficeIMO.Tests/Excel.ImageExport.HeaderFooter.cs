using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using System.Text;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelImageExportHeaderFooterTests {
        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRendersPlainHeaderFooterText() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerCenter: "Confidential", footerRight: "Draft");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Equal(2, results.Count);
            Assert.All(results, result => Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported));
            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.Contains(">Confidential<", svg);
            Assert.Contains(">Draft<", svg);
            Assert.Contains(">A3<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRendersSupportedHeaderFooterFields() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerCenter: "Sheet &A", footerRight: "Page &[Page] of &N && draft");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Equal(2, results.Count);
            Assert.All(results, result => Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported));
            string firstSvg = Encoding.UTF8.GetString(results[0].Bytes);
            string secondSvg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.Contains(">Sheet Report<", secondSvg);
            Assert.Contains(">Page 1 of 2 &amp; draft<", firstSvg);
            Assert.Contains(">Page 2 of 2 &amp; draft<", secondSvg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedPngExportAddsHeaderFooterBands() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerLeft: "Prepared", footerCenter: "Internal");
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult composed = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];
            OfficeImageExportResult bodyOnly = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                Range = "A3:D4",
                ShowGridlines = false
            });

            Assert.DoesNotContain(composed.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            OfficeImageInfo composedInfo = OfficeImageReader.Identify(composed.Bytes);
            OfficeImageInfo bodyInfo = OfficeImageReader.Identify(bodyOnly.Bytes);
            Assert.Equal(bodyInfo.Width, composedInfo.Width);
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
