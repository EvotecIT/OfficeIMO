using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using System.Text;
using Xunit;
using X = DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeIMO.Tests {
    public class ExcelImageExportPrintAreaTests {
        [Fact]
        public void ExcelWorkbookRejectsPredictableOutputCountBeforeConsumerSideEffects() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            document.AddWorksheet("One").CellValue(1, 1, "one");
            document.AddWorksheet("Two").CellValue(1, 1, "two");
            int consumed = 0;

            OfficeImageExportBatchLimitException exception = Assert.Throws<OfficeImageExportBatchLimitException>(() =>
                document.ExportImages(
                    OfficeImageExportFormat.Png,
                    _ => consumed++,
                    new ExcelImageExportOptions { MaximumOutputCount = 1 }));

            Assert.Equal(0, consumed);
            Assert.Equal(2, exception.Actual);
            Assert.Equal(1, exception.Maximum);
        }

        [Fact]
        public void ExcelWorksheet_ImageExportUsesPrintAreaWhenRequested() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            sheet.CellValue(1, 1, "Outside");
            sheet.CellValue(2, 2, "North");
            sheet.CellValue(2, 3, 10);
            sheet.CellValue(3, 2, "South");
            sheet.CellValue(3, 3, 20);
            document.SetPrintArea(sheet, "B2:C3", save: false);

            ExcelRangeVisualInfo snapshot = sheet.CreateVisualSnapshot(new ExcelImageExportOptions {
                UsePrintArea = true,
                ShowGridlines = false
            });
            OfficeImageExportResult png = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                UsePrintArea = true,
                ShowGridlines = false
            });

            Assert.Equal("B2:C3", snapshot.Range);
            Assert.Equal(new[] { 2, 3 }, snapshot.Rows.Select(row => row.Index).ToArray());
            Assert.Equal(new[] { 2, 3 }, snapshot.Columns.Select(column => column.Index).ToArray());
            Assert.DoesNotContain(snapshot.Cells, cell => cell.Text == "Outside");
            Assert.Contains(snapshot.Cells, cell => cell.Text == "North");
            Assert.Contains(snapshot.Cells, cell => cell.Text == "South");
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code.StartsWith("excel-print-area-", StringComparison.Ordinal));
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code.StartsWith("excel-print-area-", StringComparison.Ordinal));
            Assert.True(OfficeImageReader.Identify(png.Bytes).Width > 0);
        }

        [Fact]
        public void ExcelWorksheet_ImageExportExplicitRangeOverridesPrintArea() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            sheet.CellValue(1, 1, "Print");
            sheet.CellValue(3, 3, "Explicit");
            document.SetPrintArea(sheet, "A1:A1", save: false);

            ExcelRangeVisualInfo snapshot = sheet.CreateVisualSnapshot(new ExcelImageExportOptions {
                Range = "C3:C3",
                UsePrintArea = true
            });

            Assert.Equal("C3:C3", snapshot.Range);
            Assert.Single(snapshot.Cells);
            Assert.Equal("Explicit", snapshot.Cells[0].Text);
            Assert.DoesNotContain(snapshot.Diagnostics, diagnostic => diagnostic.Code.StartsWith("excel-print-area-", StringComparison.Ordinal));
        }

        [Fact]
        public void ExcelWorksheet_ImageExportReportsMissingPrintAreaFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(2, 2, "Value");

            OfficeImageExportResult result = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                UsePrintArea = true
            });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintAreaMissing);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, diagnostic.Severity);
            Assert.Equal("Report!_xlnm.Print_Area", diagnostic.Source);
        }

        [Fact]
        public void ExcelWorksheet_ImageExportReportsUnsupportedMultiAreaPrintAreaFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Used");
                sheet.CellValue(2, 2, "First");
                sheet.CellValue(2, 4, "Second");
                document.Save();
            }

            AddMultiAreaPrintArea(filePath);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loaded.Sheets[0];
            ExcelRangeVisualInfo snapshot = loadedSheet.CreateVisualSnapshot(new ExcelImageExportOptions {
                UsePrintArea = true
            });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(snapshot.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintAreaMultipleAreasUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Report!_xlnm.Print_Area", diagnostic.Source);
            Assert.Contains(snapshot.Cells, cell => cell.Text == "First");
            Assert.Contains(snapshot.Cells, cell => cell.Text == "Second");
        }

        [Fact]
        public void ExcelWorksheet_ExportImagesSplitsMultiAreaPrintAreaIntoImageResults() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(1, 1, "Used");
                sheet.CellValue(2, 2, "First");
                sheet.CellValue(2, 4, "Second");
                document.Save();
            }

            AddMultiAreaPrintArea(filePath);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loaded.Sheets[0];
            IReadOnlyList<OfficeImageExportResult> results = loadedSheet.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                UsePrintArea = true,
                ShowGridlines = false
            });
            Assert.Equal(2, results.Count);
            Assert.Equal(new int?[] { 0, 1 }, results.Select(result => result.SequenceIndex));
            Assert.All(results, result => Assert.Equal(2, result.SequenceCount));
            Assert.Equal("Report!B2:B2", results[0].Source);
            Assert.Equal("Report!D2:D2", results[1].Source);
            Assert.All(results, result => {
                OfficeImageExportDiagnostic diagnostic = Assert.Single(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintAreaMultipleAreasSplit);
                Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, diagnostic.Severity);
                Assert.Equal("Report!_xlnm.Print_Area", diagnostic.Source);
                Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintAreaMultipleAreasUnsupported);
                Assert.True(OfficeImageReader.Identify(result.Bytes).Width > 0);
            });
        }

        [Fact]
        public void ExcelWorksheet_ManualPageBreakBudgetAlsoBoundsUnsplitPrintAreas() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(2, 2, "First");
                sheet.CellValue(2, 4, "Second");
                document.Save();
            }

            AddMultiAreaPrintArea(filePath);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loaded.Sheets[0];
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                loadedSheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelImageExportOptions {
                    UsePrintArea = true,
                    SplitByManualPageBreaks = true,
                    MaximumPageBreakImages = 1
                }));

            Assert.Contains("aggregate result limit", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelWorksheet_PrintAreaResultBudgetWinsBeforeUnboundedNormalization() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                document.AddWorksheet("Report").CellValue(1, 1, "safe");
                document.Save();
            }

            string printArea = string.Join(",", Enumerable.Range(1, 1_000)
                .Select(row => "'Report'!$A$" + row + ":$A$" + row)) + ",not-a-range";
            SetPrintAreaDefinition(filePath, printArea);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
                loaded.Sheets[0].ExportImages(OfficeImageExportFormat.Svg, new ExcelImageExportOptions {
                    UsePrintArea = true,
                    SplitByManualPageBreaks = true,
                    MaximumPageBreakImages = 2
                }));

            Assert.Contains("aggregate result limit of 2", exception.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelWorkbook_ImageExportCanUseWorksheetPrintAreas() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet summary = document.AddWorksheet("Summary");
            summary.CellValue(1, 1, "Print");
            summary.CellValue(4, 4, "Outside");
            ExcelSheet details = document.AddWorksheet("Details");
            details.CellValue(1, 1, "No print area");
            document.SetPrintArea(summary, "A1:A1", save: false);
            Assert.Equal("$A$1", summary.GetPrintArea());
            OfficeImageExportResult directSummary = summary.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                UsePrintArea = true
            });
            Assert.Equal("Summary!A1:A1", directSummary.Source);

            IReadOnlyList<OfficeImageExportResult> results = document.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                UsePrintArea = true
            });

            Assert.Equal(2, results.Count);
            Assert.Equal("Summary!A1:A1", results[0].Source);
            Assert.DoesNotContain(results[0].Diagnostics, diagnostic => diagnostic.Code.StartsWith("excel-print-area-", StringComparison.Ordinal));
            Assert.Equal("Details!A1:A1", results[1].Source);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintAreaMissing);
            Assert.Equal("Details!_xlnm.Print_Area", diagnostic.Source);
        }

        [Fact]
        public void ExcelWorkbook_SaveAsImagesKeepsMultiAreaPrintAreaOutputsDistinct() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                sheet.CellValue(2, 2, "First");
                sheet.CellValue(2, 4, "Second");
                document.Save();
            }

            AddMultiAreaPrintArea(filePath);

            string folderPath = Path.Combine(Path.GetTempPath(), "OfficeIMO.Excel.Images." + Guid.NewGuid().ToString("N"));
            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            OfficeIMO.OfficeConversionReport results = loaded.SaveAsImages(folderPath, new ExcelImageExportOptions {
                UsePrintArea = true,
                ShowGridlines = false
            });

            Assert.Equal(2, results.ItemsWritten);
            Assert.True(File.Exists(Path.Combine(folderPath, "Report.png")));
            Assert.True(File.Exists(Path.Combine(folderPath, "Report-2.png")));
        }

        [Fact]
        public void ExcelWorkbook_SaveAsImagesKeepsGeneratedFileNamesUniqueWhenSheetNamesOverlapPageSuffixes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet report = document.AddWorksheet("Report");
            FillPageBreakGrid(report);
            report.AddManualRowPageBreak(2, save: false);
            ExcelSheet reportTwo = document.AddWorksheet("Report-2");
            reportTwo.CellValue(1, 1, "Overlapping sheet name");

            string folderPath = Path.Combine(Path.GetTempPath(), "OfficeIMO.Excel.Images." + Guid.NewGuid().ToString("N"));
            OfficeIMO.OfficeConversionReport results = document.SaveAsImages(folderPath, new ExcelImageExportOptions {
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            string[] fileNames = Directory.GetFiles(folderPath, "*.png")
                .Select(file => Path.GetFileName(file)!)
                .OrderBy(fileName => fileName, StringComparer.OrdinalIgnoreCase)
                .ToArray();
            Assert.Equal(3, results.ItemsWritten);
            Assert.Equal(3, fileNames.Distinct(StringComparer.OrdinalIgnoreCase).Count());
            Assert.Contains("Report.png", fileNames);
            Assert.Contains("Report-2.png", fileNames);
            Assert.Contains("Report-2-2.png", fileNames);
        }

        [Fact]
        public void ExcelWorksheet_ExportImagesSplitsRangeByManualPageBreaks() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.AddManualRowPageBreak(2, save: false);
            sheet.AddManualColumnPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Equal(new[] {
                "Report!A1:B2",
                "Report!A3:B4",
                "Report!C1:D2",
                "Report!C3:D4"
            }, results.Select(result => result.Source).ToArray());
            Assert.All(results, result => {
                OfficeImageExportDiagnostic diagnostic = Assert.Single(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ManualPageBreaksSplit);
                Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, diagnostic.Severity);
                Assert.Equal("Report!A1:D4", diagnostic.Source);
                Assert.True(OfficeImageReader.Identify(result.Bytes).Width > 0);
            });
        }

        [Fact]
        public void ExcelWorksheet_ExportImageReportsManualPageBreaksNeedMultiOutputPath() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Equal("Report!A1:D4", result.Source);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ManualPageBreaksSingleImageUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Report!A1:D4", diagnostic.Source);
        }

        [Fact]
        public void ExcelWorkbook_ExportImagesForwardsManualPageBreakSplittingAndPageOrder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.AddManualRowPageBreak(2, save: false);
            sheet.AddManualColumnPageBreak(2, save: false);
            sheet.SetPageSetup(pageOrder: ExcelPageOrder.OverThenDown);

            IReadOnlyList<OfficeImageExportResult> results = document.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                SheetNames = new[] { "Report" },
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Equal(new[] {
                "Report!A1:B2",
                "Report!C1:D2",
                "Report!A3:B4",
                "Report!C3:D4"
            }, results.Select(result => result.Source).ToArray());
            Assert.All(results, result => Assert.Contains(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ManualPageBreaksSplit));
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedImageExportReportsRemainingPageChromeDiagnostics() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: null, lastCol: null, save: false);
            sheet.SetOrientation(OfficePageOrientation.Landscape);
            sheet.SetPageSetup(fitToWidth: 2, fitToHeight: 1);
            sheet.SetHeaderFooter(headerCenter: "Confidential", footerRight: "Printed &BConfidential");
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Equal(2, results.Count);
            Assert.All(results, result => {
                Assert.DoesNotContain(result.Diagnostics, item =>
                    item.Code == ExcelImageExportDiagnosticCodes.PrintTitlesUnsupported);
                Assert.Contains(result.Diagnostics, item =>
                    item.Code == ExcelImageExportDiagnosticCodes.PageSetupUnsupported &&
                    item.Source == "Report!pageSetup");
                Assert.DoesNotContain(result.Diagnostics, item =>
                    item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterUnsupported);
                Assert.Contains(result.Diagnostics, item =>
                    item.Code == ExcelImageExportDiagnosticCodes.HeaderFooterFormattingApproximation &&
                    item.Source == "Report!headerFooter");
                Assert.Contains(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ManualPageBreaksSplit);
                Assert.True(OfficeImageReader.Identify(result.Bytes).Width > 0);
            });
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedPngExportAppliesPageSetupCanvasForManualScale() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetOrientation(OfficePageOrientation.Landscape);
            sheet.SetMargins(OfficeLength.Inches(0.5D), OfficeLength.Inches(0.5D), OfficeLength.Inches(0.5D), OfficeLength.Inches(0.5D));
            sheet.SetPageSetup(scale: 50);
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];

            OfficeImageInfo info = OfficeImageReader.Identify(result.Bytes);
            Assert.Equal(1056, info.Width);
            Assert.Equal(816, info.Height);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupUnsupported);
            Assert.Contains(result.Diagnostics, item =>
                item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted &&
                item.Source == "Report!pageSetup");
            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.NotNull(image);
            (int x, int y) = FindFirstNonWhitePixel(image!);
            Assert.InRange(x, 48, 180);
            Assert.InRange(y, 48, 140);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportAppliesPageSetupCanvasForManualScale() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetMargins(OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.5D), OfficeLength.Inches(0.5D));
            sheet.SetPageSetup(scale: 75);
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];

            string svg = Encoding.UTF8.GetString(result.Bytes);
            Assert.Equal(816, result.Width);
            Assert.Equal(1056, result.Height);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupUnsupported);
            Assert.Contains(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted);
            Assert.Contains("width=\"816\"", svg);
            Assert.Contains("height=\"1056\"", svg);
            Assert.Contains("<svg x=\"24\" y=\"48\"", svg);
            Assert.Contains("transform=\"scale(0.75)\"", svg);
            Assert.Contains(">A3<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportAppliesFitToWidthScaling() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            for (int column = 1; column <= 4; column++) {
                sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(40D));
            }

            sheet.SetMargins(OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D));
            sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0, paperSize: ExcelPaperSize.Letter);
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];

            string svg = Encoding.UTF8.GetString(result.Bytes);
            Assert.Equal(816, result.Width);
            Assert.Equal(1056, result.Height);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupUnsupported);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted);
            Assert.Contains("<svg x=\"24\" y=\"24\" width=\"768\"", svg);
            Assert.Contains(">A3<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedPngExportAppliesFitToWidthScaling() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.CellBackground(3, 4, "#00FF00");
            for (int column = 1; column <= 4; column++) {
                sheet.SetColumnWidth(column, ExcelColumnWidth.Characters(40D));
            }

            sheet.SetMargins(OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D));
            sheet.SetPageSetup(fitToWidth: 1, fitToHeight: 0, paperSize: ExcelPaperSize.Letter);
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];

            Assert.Equal(816, result.Width);
            Assert.Equal(1056, result.Height);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupUnsupported);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted);
            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.NotNull(image);
            OfficeColor fitPixel = image!.GetPixel(760, 30);
            Assert.True(fitPixel.G > 180 && fitPixel.R < 80 && fitPixel.B < 80, "Expected the far-right filled cell to be visible after fit-to-width scaling.");
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportAppliesFitToHeightScaling() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetRowHeight(3, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(800D)));
            sheet.SetRowHeight(4, ExcelColumnWidth.Length(OfficeIMO.OfficeLength.Points(800D)));
            sheet.SetMargins(OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D));
            sheet.SetPageSetup(fitToWidth: 0, fitToHeight: 1, paperSize: ExcelPaperSize.Letter);
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];

            string svg = Encoding.UTF8.GetString(result.Bytes);
            Assert.Equal(816, result.Width);
            Assert.Equal(1056, result.Height);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupUnsupported);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted);
            Assert.Contains("<svg x=\"24\" y=\"24\"", svg);
            Assert.Contains("height=\"1008\"", svg);
            Assert.Contains(">A3<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedPngExportAppliesConfiguredPaperSize() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetMargins(OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.5D), OfficeLength.Inches(0.5D));
            sheet.SetPageSetup(scale: 100, paperSize: ExcelPaperSize.A4);
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];

            OfficeImageInfo info = OfficeImageReader.Identify(result.Bytes);
            Assert.Equal(794, info.Width);
            Assert.Equal(1123, info.Height);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeUnsupported);
            Assert.True(OfficePngReader.TryDecode(result.Bytes, out OfficeRasterImage? image));
            Assert.NotNull(image);
            (int x, int y) = FindFirstNonWhitePixel(image!);
            Assert.InRange(x, 24, 120);
            Assert.InRange(y, 48, 140);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportAppliesLandscapePaperSize() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            sheet.SetOrientation(OfficePageOrientation.Landscape);
            sheet.SetMargins(OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D));
            sheet.SetPageSetup(scale: 100, paperSize: ExcelPaperSize.Legal);
            sheet.AddManualRowPageBreak(2, save: false);

            OfficeImageExportResult result = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            })[1];

            string svg = Encoding.UTF8.GetString(result.Bytes);
            Assert.Equal(1344, result.Width);
            Assert.Equal(816, result.Height);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeUnsupported);
            Assert.Contains("width=\"1344\"", svg);
            Assert.Contains("height=\"816\"", svg);
            Assert.Contains("<svg x=\"24\" y=\"24\"", svg);
            Assert.Contains(">A3<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedPngExportDiagnosesUnsupportedPaperSizeCode() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Report");
                FillPageBreakGrid(sheet);
                sheet.SetMargins(OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D), OfficeLength.Inches(0.25D));
                sheet.SetPageSetup(scale: 100);
                sheet.AddManualRowPageBreak(2, save: false);
                document.Save();
            }

            SetFirstWorksheetPaperSizeCode(filePath, 999U);

            OfficeImageExportResult result;
            using (ExcelDocument document = ExcelDocument.Load(filePath, new OfficeIMO.Excel.ExcelLoadOptions { AccessMode = OfficeIMO.DocumentAccessMode.ReadOnly })) {
                ExcelSheet sheet = document.GetSheet("Report");
                result = sheet.ExportImages(OfficeImageExportFormat.Png, new ExcelImageExportOptions {
                    Range = "A1:D4",
                    SplitByManualPageBreaks = true,
                    ShowGridlines = false
                })[1];
            }

            OfficeImageInfo info = OfficeImageReader.Identify(result.Bytes);
            Assert.Equal(816, info.Width);
            Assert.Equal(1056, info.Height);
            Assert.DoesNotContain(result.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeDefaulted);
            Assert.Contains(result.Diagnostics, item =>
                item.Code == ExcelImageExportDiagnosticCodes.PageSetupPaperSizeUnsupported &&
                item.Source == "Report!pageSetup");
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgExportRepeatsPrintTitleRows() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: null, lastCol: null, save: false);
            sheet.AddManualRowPageBreak(2, save: false);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(OfficeImageExportFormat.Svg, new ExcelImageExportOptions {
                Range = "A1:D4",
                SplitByManualPageBreaks = true,
                ShowGridlines = false
            });

            Assert.Equal(2, results.Count);
            Assert.Equal("Report!A3:D4", results[1].Source);
            Assert.DoesNotContain(results[1].Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.PrintTitlesUnsupported);
            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.Contains(">A1<", svg);
            Assert.Contains(">A3<", svg);
        }

        [Fact]
        public void ExcelWorksheet_PageSlicedSvgFitAppliesAfterPrintTitlesAndPageCanvasComposition() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("Report");
            FillPageBreakGrid(sheet);
            document.SetPrintTitles(sheet, firstRow: 1, lastRow: 1, firstCol: null, lastCol: null, save: false);
            sheet.AddManualRowPageBreak(2, save: false);
            sheet.SetPageSetup(scale: 100, paperSize: ExcelPaperSize.Letter);

            IReadOnlyList<OfficeImageExportResult> results = sheet.ExportImages(
                OfficeImageExportFormat.Svg,
                new ExcelImageExportOptions {
                    Range = "A1:D4",
                    SplitByManualPageBreaks = true,
                    ShowGridlines = false,
                    MaximumOutputWidth = 180,
                    MaximumOutputHeight = 180
                });

            Assert.Equal(2, results.Count);
            Assert.All(results, result => {
                Assert.True(result.Width <= 180);
                Assert.True(result.Height <= 180);
                OfficeImageInfo info = OfficeImageReader.Identify(result.Bytes);
                Assert.Equal(result.Width, info.Width);
                Assert.Equal(result.Height, info.Height);
            });
            string svg = Encoding.UTF8.GetString(results[1].Bytes);
            Assert.Contains(">A1<", svg);
            Assert.Contains(">A3<", svg);
        }

        private static void FillPageBreakGrid(ExcelSheet sheet) {
            for (int row = 1; row <= 4; row++) {
                for (int column = 1; column <= 4; column++) {
                    sheet.CellValue(row, column, ExcelA1.CellReference(row, column));
                }
            }
        }

        private static (int X, int Y) FindFirstNonWhitePixel(OfficeRasterImage image) {
            int left = image.Width, top = image.Height;
            for (int y = 0; y < image.Height; y++) {
                for (int x = 0; x < image.Width; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.A > 0 && (pixel.R < 245 || pixel.G < 245 || pixel.B < 245)) {
                        left = Math.Min(left, x);
                        top = Math.Min(top, y);
                    }
                }
            }

            if (left == image.Width) throw new InvalidOperationException("Expected at least one visible non-white pixel.");
            return (left, top);
        }

        private static void SetFirstWorksheetPaperSizeCode(string filePath, uint paperSizeCode) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            X.Worksheet worksheet = worksheetPart.Worksheet;
            X.PageSetup? pageSetup = worksheet.GetFirstChild<X.PageSetup>();
            if (pageSetup == null) {
                pageSetup = new X.PageSetup();
                worksheet.Append(pageSetup);
            }

            pageSetup.PaperSize = paperSizeCode;
            worksheet.Save();
        }

        private static void AddMultiAreaPrintArea(string filePath) {
            SetPrintAreaDefinition(filePath, "'Report'!$B$2:$B$2,'Report'!$D$2:$D$2");
        }

        private static void SetPrintAreaDefinition(string filePath, string definition) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorkbookPart? workbookPart = spreadsheet.WorkbookPart;
            Assert.NotNull(workbookPart);
            X.Workbook? workbook = workbookPart!.Workbook;
            Assert.NotNull(workbook);
            workbook!.DefinedNames ??= new X.DefinedNames();
            workbook.DefinedNames.Elements<X.DefinedName>()
                .Where(name => string.Equals(name.Name?.Value, "_xlnm.Print_Area", StringComparison.OrdinalIgnoreCase))
                .ToList()
                .ForEach(name => name.Remove());
            workbook.DefinedNames.Append(new X.DefinedName {
                Name = "_xlnm.Print_Area",
                LocalSheetId = 0U,
                Text = definition
            });
            workbook.Save();
        }
    }
}
