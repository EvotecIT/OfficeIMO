using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using System.Globalization;
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
        public void ExcelWorksheet_PageSlicedSvgExportRendersSupportedWorkbookFileFields() {
            string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO-HeaderFooterFields");
            Directory.CreateDirectory(directory);
            string filePath = Path.Combine(directory, "FieldWorkbook.xlsx");
            if (File.Exists(filePath)) {
                File.Delete(filePath);
            }

            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerLeft: "File &F", headerRight: "Path &[Path]", footerRight: "&[File]");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            string expectedPathPrefix = directory.EndsWith(Path.DirectorySeparatorChar.ToString(), StringComparison.Ordinal)
                ? directory
                : directory + Path.DirectorySeparatorChar;
            Assert.DoesNotContain(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.Contains(">File FieldWorkbook.xlsx<", svg);
            Assert.Contains(">FieldWorkbook.xlsx<", svg);
            Assert.Contains(">Path " + expectedPathPrefix, svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRendersSupportedDateTimeHeaderFooterFields() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            DateTime headerFooterDateTime = new DateTime(2026, 6, 23, 14, 35, 0);
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerLeft: "Date &D", headerCenter: "Time &[Time]", footerRight: "Printed &[Date] &T");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false,
                HeaderFooterDateTime = headerFooterDateTime
            });

            string expectedDate = headerFooterDateTime.ToString("d", CultureInfo.CurrentCulture);
            string expectedTime = headerFooterDateTime.ToString("t", CultureInfo.CurrentCulture);
            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.DoesNotContain(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.Contains(">Date " + expectedDate + "<", svg);
            Assert.Contains(">Time " + expectedTime + "<", svg);
            Assert.Contains(">Printed " + expectedDate + " " + expectedTime + "<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRendersFirstEvenAndOddHeaderFooterVariants() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Report");
            FillPageBreakGrid(sheet, rows: 6);
            sheet.SetHeaderFooter(headerCenter: "Odd &P", footerRight: "Odd Footer");
            sheet.SetFirstPageHeaderFooter(headerCenter: "First &P", footerRight: "First Footer");
            sheet.SetEvenPageHeaderFooter(headerCenter: "Even &P", footerRight: "Even Footer");
            sheet.AddManualRowPageBreak(2, save: false);
            sheet.AddManualRowPageBreak(4, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:D6",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Equal(3, results.Count);
            Assert.All(results, result => Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported));
            string firstSvg = Encoding.UTF8.GetString(results[0].Bytes);
            string secondSvg = Encoding.UTF8.GetString(results[1].Bytes);
            string thirdSvg = Encoding.UTF8.GetString(results[2].Bytes);
            Assert.Contains(">First 1<", firstSvg);
            Assert.Contains(">First Footer<", firstSvg);
            Assert.DoesNotContain(">Odd 1<", firstSvg);
            Assert.DoesNotContain(">Even 1<", firstSvg);
            Assert.Contains(">Even 2<", secondSvg);
            Assert.Contains(">Even Footer<", secondSvg);
            Assert.DoesNotContain(">Odd 2<", secondSvg);
            Assert.Contains(">Odd 3<", thirdSvg);
            Assert.Contains(">Odd Footer<", thirdSvg);
            Assert.DoesNotContain(">Even 3<", thirdSvg);
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

        private static void FillPageBreakGrid(ExcelSheet sheet, int rows = 4) {
            for (int row = 1; row <= rows; row++) {
                for (int column = 1; column <= 4; column++) {
                    sheet.CellValue(row, column, A1.CellReference(row, column));
                }
            }
        }
    }
}
