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
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerCenter: "Confidential", footerRight: "Draft");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:AZ4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Equal(2, results.Count);
            Assert.All(results, result => Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported));
            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.Contains("font-family=\"Calibri, Arial, sans-serif\"", svg);
            Assert.Contains(">Confidential<", svg);
            Assert.Contains(">Draft<", svg);
            Assert.Contains(">A3<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRendersSupportedHeaderFooterFields() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerCenter: "Sheet &A", footerRight: "Page &[Page] of &N && draft");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:P4",
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
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerLeft: "File &F", headerRight: "Path &[Path]", footerRight: "&[File]");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:AZ4",
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
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerLeft: "Date &D", headerCenter: "Time &[Time]", footerRight: "Printed &[Date] &T");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:H4",
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
        public void ExcelWorksheet_PageSlicedSvgExportRendersBasicFormattedHeaderFooterText() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerCenter: "&B&I&UStyled &A", footerRight: "Printed &BConfidential");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:P4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.DoesNotContain(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.Contains(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation);
            Assert.Contains("font-family=\"Calibri, Arial, sans-serif\"", svg);
            Assert.Contains("font-weight=\"700\"", svg);
            Assert.Contains("font-style=\"italic\"", svg);
            Assert.Contains("text-decoration=\"underline\"", svg);
            Assert.Contains(">Styled Report<", svg);
            Assert.Contains(">Printed ", svg);
            Assert.Contains(">Confidential<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRendersColorAndSizeHeaderFooterFormatting() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerCenter: "&KFF0000&14Red &A");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:P4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.DoesNotContain(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.Contains(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation);
            Assert.Contains("font-size=\"14\"", svg);
            Assert.Contains("fill=\"#FF0000\"", svg);
            Assert.Contains(">Red Report<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRendersStrikethroughHeaderFooterFormatting() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerCenter: "&SStrike Header");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:P4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.DoesNotContain(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.Contains(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation);
            Assert.Contains("text-decoration=\"line-through\"", svg);
            Assert.Contains("Strike Header", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRendersFontFamilyHeaderFooterFormatting() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerCenter: "&\"Aptos,Bold Italic\"Font Header");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.DoesNotContain(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.Contains(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation);
            Assert.Contains("font-family=\"Aptos\"", svg);
            Assert.Contains("font-weight=\"700\"", svg);
            Assert.Contains("font-style=\"italic\"", svg);
            Assert.Contains(">Font Header<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportReportsUnresolvedHeaderFooterFontFamilyFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerCenter: "&\"OfficeIMO Missing Header Font,Bold\"Font Header");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterFontFamilyFallback);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Report!headerFooter", diagnostic.Source);
            Assert.Contains("OfficeIMO Missing Header Font", diagnostic.Message);
            Assert.Contains(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation);
            Assert.Contains("font-family=\"OfficeIMO Missing Header Font\"", Encoding.UTF8.GetString(results[1].Bytes));
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportKeepsMalformedHeaderFooterFontFamilyDiagnosed() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerCenter: "&\"AptosFont Header");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Contains(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.DoesNotContain(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation);
            Assert.DoesNotContain("Font Header", Encoding.UTF8.GetString(results[1].Bytes));
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportClipsHeaderFooterTextZones() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            DateTime headerFooterDateTime = new DateTime(2026, 6, 23, 14, 35, 0);
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderFooter(headerLeft: "Date &D", headerCenter: "Time &[Time]", footerRight: "Printed &[Date] &T");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false,
                HeaderFooterDateTime = headerFooterDateTime
            });

            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.DoesNotContain(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.Contains("id=\"xl-header-footer-header-left\"", svg);
            Assert.Contains("id=\"xl-header-footer-header-center\"", svg);
            Assert.Contains("id=\"xl-header-footer-footer-right\"", svg);
            Assert.Contains("clip-path=\"url(#xl-header-footer-header-left)\"", svg);
            Assert.Contains("clip-path=\"url(#xl-header-footer-header-center)\"", svg);
            Assert.Contains("clip-path=\"url(#xl-header-footer-footer-right)\"", svg);
            Assert.Contains(">Date ", svg);
            Assert.Contains(">Time ", svg);
            Assert.Contains(">Printed ", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRendersFirstEvenAndOddHeaderFooterVariants() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
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
            ExcelSheet sheet = document.AddWorksheet("Report");
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

        [Fact]
        public void ExcelWorksheet_PageSlicedImageExportKeepsHeaderFooterInsidePageSetupCanvas() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetPaperSize(ExcelPaperSize.Letter);
            sheet.SetHeaderFooter(headerCenter: "Prepared", footerCenter: "Internal");
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];

            Assert.Equal(OfficePageSizes.Letter.ToPixelHeight(96D), result.Height);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedImageExportKeepsHeaderFooterInsidePageSetupCanvasWithPrintTitles() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetPaperSize(ExcelPaperSize.Letter);
            sheet.SetHeaderFooter(headerCenter: "Prepared", footerCenter: "Internal");
            document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: null, lastCol: null, save: false);
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];

            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintTitlesUnsupported);
            Assert.Equal(OfficePageSizes.Letter.ToPixelHeight(96D), result.Height);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedImageExportRendersSupportedHeaderFooterPngImages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            byte[] logo = CreateHeaderFooterLogoPng();
            sheet.SetHeaderImage(HeaderFooterPosition.Center, logo, "image/png", widthPoints: 36D, heightPoints: 16D);
            sheet.SetFooterImage(HeaderFooterPosition.Right, logo, "image/png", widthPoints: 24D, heightPoints: 12D);
            sheet.AddManualRowPageBreak(2, save: false);

            var options = new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            };
            OfficeImageExportResult png = sheet.ExportImages(OfficeImageExportFormat.Png, options)[1];
            OfficeImageExportResult svg = sheet.ExportImages(OfficeImageExportFormat.Svg, options)[1];
            string svgText = Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.Contains(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterImageApproximation);
            Assert.Contains(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterImageApproximation);
            Assert.Contains("<image", svgText);
            Assert.Contains("data:image/png;base64,", svgText);
            Assert.DoesNotContain("&G", svgText);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.NotNull(image);
            Assert.True(CountRedHeaderPixels(image!) >= 30);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedImageExportRendersHeaderFooterBmpImagesThroughSharedDecoder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderImage(HeaderFooterPosition.Center, CreateHeaderFooterLogoBmp(), "image/bmp", widthPoints: 36D, heightPoints: 16D);
            sheet.AddManualRowPageBreak(2, save: false);

            var options = new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            };
            OfficeImageExportResult png = sheet.ExportImages(OfficeImageExportFormat.Png, options)[1];
            OfficeImageExportResult svg = sheet.ExportImages(OfficeImageExportFormat.Svg, options)[1];
            string svgText = Encoding.UTF8.GetString(svg.Bytes);

            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.DoesNotContain(svg.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.Contains(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterImageApproximation);
            Assert.Contains(svg.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterImageApproximation);
            Assert.Contains("data:image/png;base64,", svgText);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? image));
            Assert.NotNull(image);
            Assert.True(CountRedHeaderPixels(image!) >= 30);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedPngExportUsesVisibleHeaderFooterImageFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderImage(HeaderFooterPosition.Center, new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 }, "image/gif", widthPoints: 36D, heightPoints: 16D);
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];

            OfficeImageExportDiagnostic diagnostic = Assert.Single(result.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Report!headerFooter", diagnostic.Source);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
            Assert.True(OfficeImageReader.Identify(result.Bytes).Width > 0);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedPngExportUsesCallerCodecForHeaderImage() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetHeaderImage(
                HeaderFooterPosition.Center,
                new byte[] { 0x47, 0x49, 0x46, 0x38, 0x39, 0x61 },
                "image/gif",
                widthPoints: 36D,
                heightPoints: 16D);
            sheet.AddManualRowPageBreak(2, save: false);
            var codec = new HeaderImageCodec();

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false,
                ImageCodec = codec
            })[1];

            Assert.Equal(2, codec.DecodeCalls);
            Assert.Contains(
                result.Diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec);
            Assert.DoesNotContain(
                result.Diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
        }

        private static void FillPageBreakGrid(ExcelSheet sheet, int rows = 4) {
            for (int row = 1; row <= rows; row++) {
                for (int column = 1; column <= 4; column++) {
                    sheet.CellValue(row, column, A1.CellReference(row, column));
                }
            }
        }

        private sealed class HeaderImageCodec : IOfficeRasterImageCodec {
            internal int DecodeCalls { get; private set; }

            public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
                DecodeCalls++;
                image = new OfficeRasterImage(4, 2, OfficeColor.FromRgb(220, 30, 30));
                return true;
            }
        }

        private static byte[] CreateHeaderFooterLogoPng() {
            OfficeRasterImage image = new OfficeRasterImage(24, 12, OfficeColor.Transparent);
            var canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 24, 12, OfficeColor.FromRgb(220, 38, 38));
            canvas.FillRectangle(4, 3, 16, 6, OfficeColor.FromRgb(255, 255, 255));
            return OfficePngWriter.Encode(image, OfficePngCompression.Stored);
        }

        private static byte[] CreateHeaderFooterLogoBmp() {
            OfficeColor red = OfficeColor.FromRgb(220, 38, 38);
            OfficeColor white = OfficeColor.White;
            var pixels = new List<OfficeColor>(24 * 12);
            for (int y = 0; y < 12; y++) {
                for (int x = 0; x < 24; x++) {
                    pixels.Add(x >= 4 && x < 20 && y >= 3 && y < 9 ? white : red);
                }
            }

            return CreateBmp24(24, 12, pixels);
        }

        private static byte[] CreateBmp24(int width, int height, IReadOnlyList<OfficeColor> pixels) {
            int rowStride = ((width * 24) + 31) / 32 * 4;
            int pixelOffset = 54;
            byte[] bytes = new byte[pixelOffset + (rowStride * height)];
            bytes[0] = (byte)'B';
            bytes[1] = (byte)'M';
            WriteInt32LittleEndian(bytes, 2, bytes.Length);
            WriteInt32LittleEndian(bytes, 10, pixelOffset);
            WriteInt32LittleEndian(bytes, 14, 40);
            WriteInt32LittleEndian(bytes, 18, width);
            WriteInt32LittleEndian(bytes, 22, height);
            WriteUInt16LittleEndian(bytes, 26, 1);
            WriteUInt16LittleEndian(bytes, 28, 24);

            for (int y = 0; y < height; y++) {
                int sourceY = height - 1 - y;
                int rowOffset = pixelOffset + (y * rowStride);
                for (int x = 0; x < width; x++) {
                    OfficeColor color = pixels[(sourceY * width) + x];
                    int offset = rowOffset + (x * 3);
                    bytes[offset] = color.B;
                    bytes[offset + 1] = color.G;
                    bytes[offset + 2] = color.R;
                }
            }

            return bytes;
        }

        private static void WriteInt32LittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
            bytes[offset + 2] = (byte)(value >> 16);
            bytes[offset + 3] = (byte)(value >> 24);
        }

        private static void WriteUInt16LittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private static int CountRedHeaderPixels(OfficeRasterImage image) {
            int count = 0;
            int maxY = Math.Min(image.Height, 60);
            for (int y = 0; y < maxY; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor actual = image.GetPixel(x, y);
                    if (actual.R >= 170 &&
                        actual.G <= 120 &&
                        actual.B <= 120 &&
                        actual.A > 200) {
                        count++;
                    }
                }
            }

            return count;
        }
    }
}
