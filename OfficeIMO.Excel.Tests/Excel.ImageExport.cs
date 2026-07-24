using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using System.Globalization;
using A = DocumentFormat.OpenXml.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using X = DocumentFormat.OpenXml.Spreadsheet;
using Xdr = DocumentFormat.OpenXml.Drawing.Spreadsheet;
using Xunit;

namespace OfficeIMO.Tests {
    public partial class ExcelImageExportTests {
        [Fact]
        public void ExcelRange_ExportsPngAndSvgFromVisualSnapshot() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(1, 2, "Value");
            sheet.CellValue(2, 1, "North");
            sheet.CellValue(2, 2, 42);
            sheet.Range("A1:B1").SetFillColor("D9EAF7").SetBold();
            sheet.CellAt(2, 2).SetBorder(BorderStyleValues.Thin, "111827");

            ExcelRange range = sheet.Range("A1:B2");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { Scale = 2 });
            string svg = range.ToSvg();

            Assert.Equal(4, snapshot.Cells.Count);
            Assert.Equal(OfficeImageExportFormat.Png, png.Format);
            OfficeImageInfo info = OfficeImageReader.Identify(png.Bytes);
            Assert.Equal(256, info.Width);
            Assert.Equal(80, info.Height);
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("<rect x=\"0\" y=\"0\" width=\"128\" height=\"40\" fill=\"#FFFFFF\"/>", svg, StringComparison.Ordinal);
            Assert.Contains("Name", svg, StringComparison.Ordinal);
            Assert.Contains("North", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ToImageFluentExportsScaledPng() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(1, 2, "Value");
            sheet.Range("A1:B1").SetFillColor("D9EAF7").SetBold();

            OfficeImageExportResult png = sheet.Range("A1:B1")
                .ToImage()
                .WithoutGridlines()
                .ForPreview()
                .WithScale(2D)
                .AsPng()
                .Export();

            Assert.Equal(OfficeImageExportFormat.Png, png.Format);
            OfficeImageInfo info = OfficeImageReader.Identify(png.Bytes);
            Assert.Equal(png.Width, info.Width);
            Assert.Equal(png.Height, info.Height);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ToImageUsesConfiguredFormatForBytes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Data");
            sheet.CellValue(1, 1, "Friendly Excel image export");

            byte[] png = sheet.Range("A1:A1")
                .ToImage()
                .WithoutGridlines()
                .AsPng()
                .ToBytes();
            string svg = System.Text.Encoding.UTF8.GetString(sheet.Range("A1:A1")
                .ToImage()
                .WithoutGridlines()
                .AsSvg()
                .ToBytes());

            Assert.Equal(new byte[] { 0x89, 0x50, 0x4E, 0x47 }, png.Take(4).ToArray());
            Assert.Contains("<svg", svg, StringComparison.Ordinal);
            Assert.Contains("<text", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("Friendly Excel image export", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelWorkbook_ToImagesPreservesHiddenSheetOptionWhenSaving() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet visible = document.AddWorksheet("Visible");
            visible.CellValue(1, 1, "Visible");
            ExcelSheet hidden = document.AddWorksheet("Hidden");
            hidden.CellValue(1, 1, "Hidden");
            hidden.SetHidden(true);

            string folder = Path.Combine(Path.GetTempPath(), "officeimo-hidden-sheets-" + Guid.NewGuid().ToString("N"));
            try {
                IReadOnlyList<OfficeImageExportResult> saved = document.ToImages()
                    .AsSvg()
                    .Save(folder);

                Assert.Single(saved);
                Assert.True(File.Exists(Path.Combine(folder, "Visible.svg")));
                Assert.False(File.Exists(Path.Combine(folder, "Hidden.svg")));

                string allFolder = Path.Combine(folder, "all");
                IReadOnlyList<OfficeImageExportResult> allSaved = document.ToImages()
                    .AsSvg()
                    .IncludeHiddenSheets()
                    .Save(allFolder);

                Assert.Equal(2, allSaved.Count);
                Assert.True(File.Exists(Path.Combine(allFolder, "Visible.svg")));
                Assert.True(File.Exists(Path.Combine(allFolder, "Hidden.svg")));
            } finally {
                if (Directory.Exists(folder)) {
                    Directory.Delete(folder, recursive: true);
                }
            }
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnresolvedCellFontFamilyFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Fonts");
            sheet.CellValue(1, 1, "Font fallback");
            sheet.CellAt(1, 1).SetFontName("OfficeIMO Missing Cell Font");
            sheet.SetColumnWidth(1, 18);

            OfficeImageExportResult png = sheet.Range("A1:A1").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Fonts!A1", diagnostic.Source);
            Assert.Contains("OfficeIMO Missing Cell Font", diagnostic.Message);
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsCallerScopedFontBeforePlatformFallback() {
            OfficeTrueTypeFont? font = OfficeTrueTypeFont.TryLoadDefault(out string? fontPath);
            if (font == null || string.IsNullOrWhiteSpace(fontPath)) return;

            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ScopedFont");
            sheet.CellValue(1, 1, "Scoped font");
            sheet.CellAt(1, 1).SetFontName("OfficeIMO Scoped Cell Font");
            sheet.SetColumnWidth(1, 18);
            var options = new ExcelImageExportOptions { ShowGridlines = false };
            options.Fonts.Add("OfficeIMO Scoped Cell Font", File.ReadAllBytes(fontPath));

            OfficeImageExportResult png = sheet.Range("A1:A1").ExportImage(OfficeImageExportFormat.Png, options);

            Assert.DoesNotContain(
                png.Diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
        }

        [Fact]
        public void WorkbookDirectStreamingEnforcesOneAggregateBudgetAcrossSheets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            document.AddWorksheet("One").CellValue(1, 1, "One");
            document.AddWorksheet("Two").CellValue(1, 1, "Two");
            var options = new ExcelWorkbookImageExportOptions {
                MaximumOutputCount = 1
            };
            int consumed = 0;

            OfficeImageExportBatchLimitException exception =
                Assert.Throws<OfficeImageExportBatchLimitException>(
                    () => document.ExportImages(
                        OfficeImageExportFormat.Png,
                        _ => consumed++,
                        options));

            Assert.Equal(1, consumed);
            Assert.Equal(nameof(OfficeImageExportOptions.MaximumOutputCount), exception.LimitName);
        }

        [Fact]
        public void ExcelRange_ImageExportUsesNumberFormatLiteralsAndSections() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Formats");
            sheet.Cell(1, 1, 5, numberFormat: "0 \"days\"");
            sheet.Cell(1, 2, 2, numberFormat: "0\\h");
            sheet.Cell(1, 3, -3, numberFormat: "0 \"up\";0 \"down\";\"flat\"");
            sheet.Cell(1, 4, 0, numberFormat: "0 \"up\";0 \"down\";\"flat\"");
            sheet.Cell(1, 5, -4, numberFormat: "#,##0;(#,##0)");

            ExcelRange range = sheet.Range("A1:E1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("5 days", snapshot.Cells.Single(cell => cell.Column == 1).Text);
            Assert.Equal("2h", snapshot.Cells.Single(cell => cell.Column == 2).Text);
            Assert.Equal("3 down", snapshot.Cells.Single(cell => cell.Column == 3).Text);
            Assert.Equal("flat", snapshot.Cells.Single(cell => cell.Column == 4).Text);
            Assert.Equal("(4)", snapshot.Cells.Single(cell => cell.Column == 5).Text);
            Assert.Contains("5 days", svg, StringComparison.Ordinal);
            Assert.Contains("2h", svg, StringComparison.Ordinal);
            Assert.Contains("3 down", svg, StringComparison.Ordinal);
            Assert.Contains("flat", svg, StringComparison.Ordinal);
            Assert.Contains("(4)", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            foreach (ExcelVisualCell cell in snapshot.Cells) {
                Assert.True(ContainsDarkPixel(rendered!, cell), $"Expected formatted text pixels in R{cell.Row}C{cell.Column}.");
            }
        }

        [Fact]
        public void ExcelRange_ImageExportUsesBuiltInNumberFormatIdsWithoutCustomFormatCodes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("BuiltIns");
            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 14);
            sheet.SetColumnWidth(3, 16);
            sheet.SetColumnWidth(4, 14);
            sheet.CellValue(1, 1, 0.1234);
            sheet.CellValue(1, 2, -1234);
            sheet.CellValue(1, 3, new DateTime(2026, 6, 24).ToOADate());
            sheet.CellValue(1, 4, 1.5);
            ApplyBuiltInNumberFormatId(document, sheet, "A1", 10U);
            ApplyBuiltInNumberFormatId(document, sheet, "B1", 37U);
            ApplyBuiltInNumberFormatId(document, sheet, "C1", 14U);
            ApplyBuiltInNumberFormatId(document, sheet, "D1", 46U);

            ExcelRange range = sheet.Range("A1:D1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("12.34%", snapshot.Cells.Single(cell => cell.Column == 1).Text);
            Assert.Equal("(1,234)", snapshot.Cells.Single(cell => cell.Column == 2).Text);
            Assert.Equal("6/24/2026", snapshot.Cells.Single(cell => cell.Column == 3).Text);
            Assert.Equal("36:00:00", snapshot.Cells.Single(cell => cell.Column == 4).Text);
            Assert.Contains("12.34%", svg, StringComparison.Ordinal);
            Assert.Contains("(1,234)", svg, StringComparison.Ordinal);
            Assert.Contains("6/24/2026", svg, StringComparison.Ordinal);
            Assert.Contains("36:00:00", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            foreach (ExcelVisualCell cell in snapshot.Cells) {
                Assert.True(ContainsDarkPixel(rendered!, cell), $"Expected built-in formatted text pixels in R{cell.Row}C{cell.Column}.");
            }
        }

        [Fact]
        public void ExcelNumberFormatDisplayFallsBackToBuiltInCodesWhenFormatCodeIsMissing() {
            Assert.Equal("12.34%", ExcelNumberFormatDisplay.FormatNumericText(0.1234D, 10U, null, "0.1234"));
            Assert.Equal("(1,234)", ExcelNumberFormatDisplay.FormatNumericText(-1234D, 37U, null, "-1234"));
            Assert.Equal("6/24/2026", ExcelNumberFormatDisplay.FormatNumericText(new DateTime(2026, 6, 24).ToOADate(), 14U, null, "46200"));
            Assert.Equal("36:00:00", ExcelNumberFormatDisplay.FormatNumericText(1.5D, 46U, null, "1.5"));
            Assert.Equal("1/2", ExcelNumberFormatDisplay.FormatNumericText(0.5D, 12U, null, "0.5"));
            Assert.Equal("1 1/4", ExcelNumberFormatDisplay.FormatNumericText(1.25D, 12U, null, "1.25"));
            Assert.Equal("-1/2", ExcelNumberFormatDisplay.FormatNumericText(-0.5D, 12U, null, "-0.5"));
            Assert.Equal("1/10", ExcelNumberFormatDisplay.FormatNumericText(0.1D, 13U, null, "0.1"));
            Assert.Equal("1/2", ExcelNumberFormatDisplay.FormatNumericText(0.5D, 1U, "# ?/2", "0.5"));
            Assert.Equal("1073741824/2147483647", ExcelNumberFormatDisplay.FormatNumericText(0.5D, 1U, "# ?/2147483647", "0.5"));
            Assert.Equal("3221225471/2147483647", ExcelNumberFormatDisplay.FormatNumericText(1.5D, 1U, "?/2147483647", "1.5"));
            Assert.Equal("(1/2)", ExcelNumberFormatDisplay.FormatNumericText(-0.5D, 1U, "# ?/?;(# ?/?)", "-0.5"));
            Assert.Equal("1,235", ExcelNumberFormatDisplay.FormatNumericText(1234567D, 1U, "#,##0,", "1234567"));
            Assert.Equal("1 K", ExcelNumberFormatDisplay.FormatNumericText(1234D, 1U, "#,##0, \"K\"", "1234"));
            Assert.Equal("1", ExcelNumberFormatDisplay.FormatNumericText(1234567D, 1U, "#,##0,,", "1234567"));
            Assert.Equal("1", ExcelNumberFormatDisplay.FormatNumericText(1D, 1U, "0.##", "1"));
            Assert.Equal("1.5", ExcelNumberFormatDisplay.FormatNumericText(1.5D, 1U, "0.##", "1.5"));
            Assert.Equal("1,234.5", ExcelNumberFormatDisplay.FormatNumericText(1234.5D, 1U, "#,##0.##", "1234.5"));
            Assert.Equal("50 low", ExcelNumberFormatDisplay.FormatNumericText(50D, 1U, "[>=100]0 \"high\";0 \"low\"", "50"));
            Assert.Equal("150 high", ExcelNumberFormatDisplay.FormatNumericText(150D, 1U, "[>=100]0 \"high\";0 \"low\"", "150"));
            Assert.Equal("6/24/26 13:45", ExcelNumberFormatDisplay.FormatNumericText(new DateTime(2026, 6, 24, 13, 45, 0).ToOADate(), 1U, "m/d/yy h:mm", "46200.5729"));
            Assert.Equal("6/24/2026 13:45", ExcelNumberFormatDisplay.FormatNumericText(new DateTime(2026, 6, 24, 13, 45, 0).ToOADate(), 1U, "m/d/yyyy h:mm", "46200.5729"));
            Assert.Equal("6/24/2026 1:45 PM", ExcelNumberFormatDisplay.FormatNumericText(new DateTime(2026, 6, 24, 13, 45, 0).ToOADate(), 1U, "m/d/yyyy h:mm AM/PM", "46200.5729"));
            Assert.Equal("01/05/26", ExcelNumberFormatDisplay.FormatNumericText(new DateTime(2026, 1, 5).ToOADate(), 1U, "mm/dd/yy", "46027"));
            Assert.Equal("January 5, 2026", ExcelNumberFormatDisplay.FormatNumericText(new DateTime(2026, 1, 5).ToOADate(), 1U, "mmmm d, yyyy", "46027"));
            Assert.Equal("Monday, January 5", ExcelNumberFormatDisplay.FormatNumericText(new DateTime(2026, 1, 5).ToOADate(), 1U, "dddd, mmmm d", "46027"));
            Assert.Equal("46027", ExcelNumberFormatDisplay.FormatNumericText(46027D, 1U, "mmmm d \"%\"", "46027"));
            Assert.Equal("5%", ExcelNumberFormatDisplay.FormatNumericText(5D, 1U, "0\"%\"", "5"));
            Assert.Equal("500%", ExcelNumberFormatDisplay.FormatNumericText(5D, 1U, "0%", "5"));
            Assert.Equal("1234%%", ExcelNumberFormatDisplay.FormatNumericText(0.1234D, 1U, "0%%", "0.1234"));
            Assert.Equal("1.23E-02", ExcelNumberFormatDisplay.FormatNumericText(0.0123D, 1U, "0.00E-00", "0.0123"));
            Assert.Equal("ver. 1", ExcelNumberFormatDisplay.FormatNumericText(1D, 1U, "\"ver. \"0", "1"));
            Assert.Equal("v0 1", ExcelNumberFormatDisplay.FormatNumericText(1D, 1U, "\"v0 \"0", "1"));
            Assert.Equal("10", ExcelNumberFormatDisplay.FormatNumericText(1D, 1U, "0\"0\"", "1"));
            Assert.Equal("1.", ExcelNumberFormatDisplay.FormatNumericText(1D, 1U, "0\".\"", "1"));
            Assert.Equal("1.", ExcelNumberFormatDisplay.FormatNumericText(1D, 1U, "0\\.", "1"));
            Assert.Equal("1;", ExcelNumberFormatDisplay.FormatNumericText(1D, 1U, "0\\;", "1"));
            Assert.Equal("$1,234.50", ExcelNumberFormatDisplay.FormatNumericText(1234.5D, 1U, "[$$-409]#,##0.00", "1234.5"));
            Assert.Equal("\u20AC1,234.50", ExcelNumberFormatDisplay.FormatNumericText(1234.5D, 1U, "[$\u20AC-407]#,##0.00", "1234.5"));
            Assert.Equal("90:00", ExcelNumberFormatDisplay.FormatNumericText(TimeSpan.FromMinutes(90).TotalDays, 1U, "[mm]:ss", "0.0625"));
            Assert.Equal("5400", ExcelNumberFormatDisplay.FormatNumericText(TimeSpan.FromMinutes(90).TotalDays, 1U, "[ss]", "0.0625"));
            Assert.Equal("1E+20", ExcelNumberFormatDisplay.FormatNumericText(1E20D, 1U, "[ss]", "1E+20"));
            Assert.Equal(string.Empty, ExcelNumberFormatDisplay.FormatNumericText(0D, 1U, "0;-0;;@", "0"));
            Assert.Equal(string.Empty, ExcelNumberFormatDisplay.FormatNumericText(0D, 1U, "#", "0"));
            Assert.Equal(string.Empty, ExcelNumberFormatDisplay.FormatNumericText(0D, 1U, "#,###", "0"));
            Assert.Equal("0", ExcelNumberFormatDisplay.FormatNumericText(0D, 1U, "0", "0"));
        }

        [Fact]
        public void ExcelRange_ImageExportUsesExcelGeneralAlignmentByValueKind() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("GeneralAlign");
            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 14);
            sheet.SetColumnWidth(3, 16);
            sheet.CellValue(1, 1, "Label");
            sheet.CellValue(1, 2, 42);
            sheet.Cell(1, 3, new DateTime(2026, 6, 24).ToOADate(), numberFormat: "m/d/yyyy");

            ExcelRange range = sheet.Range("A1:C1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            string svg = range.ToSvg(options);

            Assert.Equal(ExcelVisualCellValueKind.Text, snapshot.Cells.Single(cell => cell.Column == 1).ValueKind);
            Assert.Equal(ExcelVisualCellValueKind.Number, snapshot.Cells.Single(cell => cell.Column == 2).ValueKind);
            Assert.Equal(ExcelVisualCellValueKind.Date, snapshot.Cells.Single(cell => cell.Column == 3).ValueKind);
            Assert.Contains("text-anchor=\"start\">Label</text>", svg, StringComparison.Ordinal);
            Assert.Contains("text-anchor=\"end\">42</text>", svg, StringComparison.Ordinal);
            Assert.Contains("text-anchor=\"end\">6/24/2026</text>", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportLaysOutMultilineCellTextThroughSharedDrawingLayout() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Wrap");
            sheet.SetColumnWidth(1, 12);
            sheet.SetRowHeight(1, 54);
            sheet.CellValue(1, 1, "Alpha\nBeta\nGamma");
            sheet.WrapCells(1, 1, 1);

            ExcelRange range = sheet.Range("A1:A1");
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped);
            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped);
            Assert.True(CountOccurrences(svg, "<text ") >= 3, "Expected multiline Excel text export to emit separate SVG text elements.");
            Assert.Contains("Alpha", svg, StringComparison.Ordinal);
            Assert.Contains("Beta", svg, StringComparison.Ordinal);
            Assert.Contains("Gamma", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(pngResult.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(ContainsDarkPixel(rendered!, range.CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false }).Cells[0]));
        }

        [Fact]
        public void ExcelRange_SvgExportDoesNotEmbedClippedCellTextSuffix() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Clip");
            sheet.SetColumnWidth(1, 8);
            sheet.SetRowHeight(1, 22);
            sheet.CellValue(1, 1, "Visible prefix SECRET-SUFFIX-MUST-NOT-LEAK");

            ExcelRange range = sheet.Range("A1:A1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);

            Assert.Contains(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "Clip!A1");
            Assert.Contains(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "Clip!A1");
            Assert.Contains("...", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("SECRET-SUFFIX-MUST-NOT-LEAK", svg, StringComparison.Ordinal);
            Assert.Contains("clip-path=\"url(#xl-text-1-1)\"", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportSpillsPlainTextIntoBlankNeighborCells() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Spill");
            sheet.SetColumnWidth(1, 6);
            sheet.SetColumnWidth(2, 24);
            sheet.SetColumnWidth(3, 8);
            sheet.SetRowHeight(1, 24);
            sheet.CellValue(1, 1, "Plain text spills");
            sheet.CellValue(1, 3, "Stop");

            ExcelRange range = sheet.Range("A1:C1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);

            ExcelVisualCell first = snapshot.Cells.Single(cell => cell.Column == 1);
            ExcelVisualCell blankNeighbor = snapshot.Cells.Single(cell => cell.Column == 2);
            double expectedWidth = first.Width + blankNeighbor.Width;
            Assert.Equal(expectedWidth, ExtractSvgClipWidth(svg, "xl-text-1-1"), precision: 2);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "Spill!A1");
            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "Spill!A1");
            Assert.Contains("Plain text spills", svg, StringComparison.Ordinal);
            Assert.Contains("Stop", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("...", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportSpillsRightAlignedTextIntoBlankLeftNeighborCells() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("RightSpill");
            sheet.SetColumnWidth(1, 24);
            sheet.SetColumnWidth(2, 6);
            sheet.SetColumnWidth(3, 8);
            sheet.SetRowHeight(1, 24);
            sheet.CellValue(1, 2, "Right aligned text spills left");
            sheet.CellValue(1, 3, "Stop");
            sheet.CellAlign(1, 2, HorizontalAlignmentValues.Right);

            ExcelRange range = sheet.Range("A1:C1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);

            ExcelVisualCell blankNeighbor = snapshot.Cells.Single(cell => cell.Column == 1);
            ExcelVisualCell source = snapshot.Cells.Single(cell => cell.Column == 2);
            double expectedWidth = blankNeighbor.Width + source.Width;
            Assert.Equal(blankNeighbor.X, ExtractSvgClipX(svg, "xl-text-1-2"), precision: 2);
            Assert.Equal(expectedWidth, ExtractSvgClipWidth(svg, "xl-text-1-2"), precision: 2);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "RightSpill!B1");
            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "RightSpill!B1");
            Assert.Contains("Right aligned text spills left", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportDoesNotSpillThroughBlankCellsCoveredByDrawings() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("OverlaySpill");
            sheet.SetColumnWidth(1, 6);
            sheet.SetColumnWidth(2, 12);
            sheet.SetColumnWidth(3, 12);
            sheet.SetRowHeight(1, 28);
            sheet.CellValue(1, 1, "Plain text stops before image overlay");
            sheet.AddImage(1, 2, CreateSolidOpaquePng(24, 18, OfficeColor.FromRgb(37, 99, 235)), "image/png", widthPixels: 24, heightPixels: 18, name: "Overlay");

            ExcelRange range = sheet.Range("A1:C1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);

            ExcelVisualCell first = snapshot.Cells.Single(cell => cell.Column == 1);
            Assert.Single(snapshot.Images);
            Assert.Equal(first.Width, ExtractSvgClipWidth(svg, "xl-text-1-1"), precision: 2);
            Assert.Contains(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "OverlaySpill!A1");
            Assert.Contains(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "OverlaySpill!A1");
            Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportAllowsSpillBehindPotentiallyTransparentCharts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartSpill");
            sheet.SetColumnWidth(1, 6);
            sheet.SetColumnWidth(2, 14);
            sheet.SetColumnWidth(3, 14);
            sheet.SetRowHeight(1, 30);
            sheet.SetRowHeight(2, 30);
            sheet.CellValue(1, 1, "Plain text stops before chart overlay");
            sheet.CellValue(3, 1, "Label");
            sheet.CellValue(3, 2, "Value");
            sheet.CellValue(4, 1, "North");
            sheet.CellValue(4, 2, 12);
            sheet.CellValue(5, 1, "South");
            sheet.CellValue(5, 2, 18);
            sheet.AddChartFromRange("A3:B5", row: 1, column: 2, widthPixels: 140, heightPixels: 70, type: ExcelChartType.ColumnClustered, title: "Overlay");

            ExcelRange range = sheet.Range("A1:C2");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);

            ExcelVisualCell first = snapshot.Cells.Single(cell => cell.Column == 1 && cell.Row == 1);
            Assert.Single(snapshot.Charts);
            Assert.True(ExtractSvgClipWidth(svg, "xl-text-1-1") > first.Width);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextOccludedByDrawing && diagnostic.Source == "ChartSpill!A1");
            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextOccludedByDrawing && diagnostic.Source == "ChartSpill!A1");
            Assert.Contains("Overlay", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportDoesNotSpillFromCellsCoveredByDrawings() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("SourceOverlay");
            sheet.SetColumnWidth(1, 10);
            sheet.SetColumnWidth(2, 16);
            sheet.SetRowHeight(1, 28);
            sheet.CellValue(1, 1, "Text hidden by image should not spill");
            sheet.AddImage(1, 1, CreateSolidOpaquePng(48, 18, OfficeColor.FromRgb(37, 99, 235)), "image/png", widthPixels: 48, heightPixels: 18, name: "SourceOverlay");

            ExcelRange range = sheet.Range("A1:B1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);

            ExcelVisualCell first = snapshot.Cells.Single(cell => cell.Column == 1);
            Assert.Single(snapshot.Images);
            Assert.Equal(first.Width, ExtractSvgClipWidth(svg, "xl-text-1-1"), precision: 2);
            Assert.Contains(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "SourceOverlay!A1");
            Assert.Contains(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "SourceOverlay!A1");
        }

        [Fact]
        public void ExcelRange_ImageExportSuppressesTextWhenDrawingCoversTextAnchor() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("AnchorOverlay");
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 24);
            sheet.CellValue(1, 1, "Covered anchor text should not show fragments");
            sheet.AddImage(1, 1, CreateSolidOpaquePng(120, 32, OfficeColor.FromRgb(37, 99, 235)), "image/png", widthPixels: 120, heightPixels: 32, name: "AnchorOverlay");

            ExcelRange range = sheet.Range("A1:A1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);

            Assert.DoesNotContain("Covered anchor text", svg, StringComparison.Ordinal);
            Assert.Contains(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextOccludedByDrawing && diagnostic.Source == "AnchorOverlay!A1");
            Assert.Contains(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextOccludedByDrawing && diagnostic.Source == "AnchorOverlay!A1");
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "AnchorOverlay!A1");
            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "AnchorOverlay!A1");
        }

        [Fact]
        public void ExcelRange_ImageExportKeepsTextUnderTransparentImageOverlays() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("TransparentOverlay");
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 24);
            sheet.CellValue(1, 1, "Visible through transparency");
            sheet.AddImage(1, 1, CreateSolidPng(120, 32, OfficeColor.Transparent), "image/png", widthPixels: 120, heightPixels: 32, name: "TransparentOverlay");

            ExcelRange range = sheet.Range("A1:A1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);

            Assert.False(Assert.Single(snapshot.Images).IsFullyOpaque);
            Assert.Contains("Visible through", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextOccludedByDrawing);
        }

        [Fact]
        public void ExcelRange_ImageExportDoesNotSpillGeneralNumericTextIntoBlankNeighbors() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("NoNumberSpill");
            sheet.SetColumnWidth(1, 6);
            sheet.SetColumnWidth(2, 24);
            sheet.SetRowHeight(1, 24);
            sheet.CellValue(1, 1, 123456789);

            ExcelRange range = sheet.Range("A1:B1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svg = System.Text.Encoding.UTF8.GetString(svgResult.Bytes);

            ExcelVisualCell first = snapshot.Cells.Single(cell => cell.Column == 1);
            Assert.Equal(ExcelVisualCellValueKind.Number, first.ValueKind);
            Assert.Equal(first.Width, ExtractSvgClipWidth(svg, "xl-text-1-1"), precision: 2);
            Assert.Contains(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "NoNumberSpill!A1");
            Assert.Contains("text-anchor=\"end\">", svg, StringComparison.Ordinal);
            Assert.Contains("...</text>", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(">123456789</text>", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsCellTextIndentAcrossSnapshotsAndRenderers() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Indent");
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 28);
            sheet.CellValue(1, 1, "Indented");
            ApplyCellTextIndentStyle(document, sheet, "A1", 3U);

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot visualSnapshot = range.CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelWorkbookSnapshot inspectionSnapshot = document.CreateInspectionSnapshot();
            ExcelCellStyleSnapshot directStyle = sheet.CellAt(1, 1).GetStyle();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            ExcelVisualCell cell = visualSnapshot.Cells[0];
            Assert.Equal(3U, cell.Style.TextIndent);
            Assert.Equal(3U, directStyle.TextIndent);
            Assert.Equal(3U, inspectionSnapshot.Worksheets[0].Cells[0].Style!.TextIndent);
            double textX = ExtractSvgTextX(svg, "Indented");
            Assert.True(textX >= cell.X + 30D, "Expected SVG text to be inset by the authored alignment indentation.");
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            (int minX, _) = DarkPixelXExtent(rendered!, cell);
            Assert.True(minX >= cell.X + 25D, $"Expected PNG text pixels to be indented. minX={minX}, cellX={cell.X}.");
        }

        [Fact]
        public void ExcelRange_ImageExportResolvesThemeTintColorsAcrossSnapshotsAndRenderers() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Theme");
            sheet.CellValue(1, 1, "Theme");
            sheet.SetColumnWidth(1, 14);
            sheet.SetRowHeight(1, 34);
            ApplyThemeBackedCellStyle(document, sheet);

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot visualSnapshot = range.CreateVisualSnapshot();
            ExcelWorkbookSnapshot inspectionSnapshot = document.CreateInspectionSnapshot();
            ExcelCellStyleSnapshot directStyle = sheet.CellAt(1, 1).GetStyle();
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            ExcelCellStyleSnapshot visualStyle = visualSnapshot.Cells[0].Style;
            ExcelCellSnapshot inspectedCell = inspectionSnapshot.Worksheets[0].Cells[0];
            Assert.Equal("FF95B3D7", visualStyle.FillColorArgb);
            Assert.Equal("FF632523", visualStyle.FontColorArgb);
            Assert.Equal("FF9BBB59", visualStyle.Border!.Top!.ColorArgb);
            Assert.Equal(visualStyle.FillColorArgb, directStyle.FillColorArgb);
            Assert.Equal(visualStyle.FontColorArgb, directStyle.FontColorArgb);
            Assert.Equal(visualStyle.FillColorArgb, inspectedCell.Style!.FillColorArgb);
            Assert.Contains("#95B3D7", svg, StringComparison.Ordinal);
            Assert.Contains("#632523", svg, StringComparison.Ordinal);
            Assert.Contains("#9BBB59", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            ExcelVisualCell renderedCell = visualSnapshot.Cells[0];
            AssertPixelNear(
                rendered!,
                (int)(renderedCell.X + renderedCell.Width - 8),
                (int)(renderedCell.Y + renderedCell.Height - 8),
                OfficeColor.FromRgb(149, 179, 215),
                tolerance: 3);
        }

        [Fact]
        public void ExcelThemeColorResolver_NormalizesInvalidSpreadsheetTintValues() {
            Assert.Equal(
                OfficeIMO.Excel.Utilities.ExcelThemeColorResolver.ApplySpreadsheetTint("FF336699", 1D),
                OfficeIMO.Excel.Utilities.ExcelThemeColorResolver.ApplySpreadsheetTint("FF336699", 2D));
            Assert.Equal(
                OfficeIMO.Excel.Utilities.ExcelThemeColorResolver.ApplySpreadsheetTint("FF336699", -1D),
                OfficeIMO.Excel.Utilities.ExcelThemeColorResolver.ApplySpreadsheetTint("FF336699", -2D));
            Assert.Equal(
                "FF336699",
                OfficeIMO.Excel.Utilities.ExcelThemeColorResolver.ApplySpreadsheetTint("FF336699", double.NaN));
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesConditionalColorScalesAndDataBarsWithDiagnostics() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Conditional");
            for (int row = 1; row <= 3; row++) {
                sheet.CellValue(row, 1, row);
                sheet.CellValue(row, 2, row);
                sheet.CellValue(row, 3, row);
            }

            sheet.SetColumnWidth(1, 12);
            sheet.SetColumnWidth(2, 12);
            sheet.SetColumnWidth(3, 12);
            sheet.AddConditionalColorScale("A1:A3", OfficeColor.Red, OfficeColor.Lime);
            sheet.AddConditionalDataBar("B1:B3", OfficeColor.Blue);
            sheet.AddConditionalIconSet("C1:C3");

            ExcelRange range = sheet.Range("A1:C3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal("FFFF0000", snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FF808000", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FF00FF00", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal(3, snapshot.ConditionalDataBars.Count);
            ExcelVisualConditionalDataBar finalBar = snapshot.ConditionalDataBars.Single(bar => bar.Row == 3 && bar.Column == 2);
            Assert.Equal("FF0000FF", finalBar.ColorArgb);
            Assert.Equal(0D, finalBar.StartRatio);
            Assert.Equal(1D, finalBar.Ratio);
            Assert.Equal(3, snapshot.ConditionalIcons.Count);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 1 && icon.Column == 3 && icon.Kind == ExcelConditionalIconKind.RedCircle);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 2 && icon.Column == 3 && icon.Kind == ExcelConditionalIconKind.YellowCircle);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 3 && icon.Column == 3 && icon.Kind == ExcelConditionalIconKind.GreenCircle);
            Assert.All(snapshot.ConditionalIcons, icon => Assert.True(icon.ShowValue));
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, diagnostic.Severity);
            Assert.Equal("Conditional!C1:C3", diagnostic.Source);
            Assert.Contains("#FF0000", svg, StringComparison.Ordinal);
            Assert.Contains("#0000FF", svg, StringComparison.Ordinal);
            Assert.Contains("#16A34A", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            ExcelVisualCell firstScaleCell = snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1);
            AssertPixelNear(
                rendered!,
                (int)(firstScaleCell.X + firstScaleCell.Width - 8),
                (int)(firstScaleCell.Y + firstScaleCell.Height - 8),
                OfficeColor.Red,
                tolerance: 3);
            OfficeColor barPixel = rendered!.GetPixel((int)(finalBar.X + finalBar.Width - 8), (int)(finalBar.Y + (finalBar.Height / 2D)));
            Assert.True(barPixel.B > 120 && barPixel.R < 40 && barPixel.G < 40, "Expected a blue data-bar pixel.");
            ExcelVisualConditionalIcon finalIcon = snapshot.ConditionalIcons.Single(icon => icon.Row == 3 && icon.Column == 3);
            Assert.True(CountGreenIconPixels(rendered!, finalIcon) > 4, "Expected visible green conditional-formatting icon pixels.");
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsConditionalDataBarThresholds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("DataBarThresholds");
                sheet.CellValue(1, 1, 0);
                sheet.CellValue(2, 1, 50);
                sheet.CellValue(3, 1, 100);
                sheet.SetColumnWidth(1, 14);
                sheet.AddConditionalDataBar("A1:A3", OfficeColor.Blue);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                DataBar dataBar = worksheet.Elements<ConditionalFormatting>().First().Elements<ConditionalFormattingRule>().First().GetFirstChild<DataBar>()!;
                ConditionalFormatValueObject[] thresholds = dataBar.Elements<ConditionalFormatValueObject>().ToArray();
                thresholds[0].Type = ConditionalFormatValueObjectValues.Number;
                thresholds[0].Val = "0";
                thresholds[1].Type = ConditionalFormatValueObjectValues.Number;
                thresholds[1].Val = "200";
                worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A3").CreateVisualSnapshot();

                ExcelVisualConditionalDataBar finalBar = snapshot.ConditionalDataBars.Single(bar => bar.Row == 3 && bar.Column == 1);
                Assert.Equal(0D, finalBar.StartRatio);
                Assert.Equal(0.5D, finalBar.Ratio, precision: 3);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportScalesDataBarsAgainstFullRuleRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("PartialDataBar");
            for (int row = 1; row <= 10; row++) {
                sheet.CellValue(row, 1, row);
            }

            sheet.AddConditionalDataBar("A1:A10", OfficeColor.Blue);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A5:A5").CreateVisualSnapshot();

            ExcelVisualConditionalDataBar bar = Assert.Single(snapshot.ConditionalDataBars);
            Assert.Equal(5, bar.Row);
            Assert.Equal(4D / 9D, bar.Ratio, precision: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportScalesColorScaleAgainstFullRuleRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("PartialColorScale");
            for (int row = 1; row <= 10; row++) {
                sheet.CellValue(row, 1, row);
            }

            sheet.AddConditionalColorScale("A1:A10", OfficeColor.Red, OfficeColor.Lime);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A5:A5").CreateVisualSnapshot();

            Assert.Equal("FF8E7100", Assert.Single(snapshot.Cells).Style.FillColorArgb);
        }

        [Fact]
        public void ExcelRange_ImageExportOmitsConditionalRulesBeyondReferenceLimit() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("OversizedConditional");
            sheet.CellValue(1, 1, 0D);
            sheet.CellValue(100_001, 1, 1000D);
            sheet.AddConditionalColorScale("A1:A100001", OfficeColor.Red, OfficeColor.Lime);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A1").CreateVisualSnapshot();

            Assert.Null(Assert.Single(snapshot.Cells).Style.FillColorArgb);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(
                snapshot.Diagnostics,
                item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalReferenceLimitExceeded);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal(OfficeImageExportLossKind.Omission, diagnostic.LossKind);
            Assert.Equal("OversizedConditional!A1:A100001", diagnostic.Source);
        }

        [Fact]
        public void ExcelRange_ImageExportBoundsAggregateConditionalRuleReferences() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("AggregateConditional");
            sheet.CellValue(1, 1, 0D);
            sheet.AddConditionalColorScale("A1:A60000", OfficeColor.Red, OfficeColor.Lime);
            sheet.AddConditionalColorScale("A1:A60000", OfficeColor.Blue, OfficeColor.White);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A1").CreateVisualSnapshot();

            OfficeImageExportDiagnostic diagnostic = Assert.Single(
                snapshot.Diagnostics,
                item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalReferenceLimitExceeded);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal(OfficeImageExportLossKind.Omission, diagnostic.LossKind);
            Assert.Equal("AggregateConditional!A1:A60000", diagnostic.Source);
            Assert.Contains("aggregate", diagnostic.Message, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExcelRange_ImageExportBoundsConditionalRuleCount() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("RuleCount");
            sheet.CellValue(1, 1, 0D);
            var conditional = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
            };
            for (int priority = 1; priority <= 4_097; priority++) {
                conditional.Append(new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.ColorScale,
                    Priority = priority
                });
            }
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            worksheet.InsertAfter(conditional, worksheet.GetFirstChild<SheetData>());

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A1").CreateVisualSnapshot();

            OfficeImageExportDiagnostic diagnostic = Assert.Single(
                snapshot.Diagnostics,
                item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalReferenceLimitExceeded);
            Assert.Contains("4096-rule", diagnostic.Message, StringComparison.Ordinal);
            Assert.Equal("RuleCount!A1:A1", diagnostic.Source);
        }

        [Fact]
        public void ExcelRange_ImageExportRetainsHigherPrecedenceRuleWithinBoundedLookahead() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("RulePriority");
            var conditional = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
            };
            conditional.Append(
                new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 100,
                    StopIfTrue = false
                },
                new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.Expression,
                    Priority = 1,
                    StopIfTrue = true
                });
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            worksheet.InsertAfter(conditional, worksheet.GetFirstChild<SheetData>());

            IReadOnlyList<ExcelConditionalFormattingInfo> retained =
                sheet.GetConditionalFormattingRules("A1", 1, out bool truncated);

            Assert.True(truncated);
            ExcelConditionalFormattingInfo rule = Assert.Single(retained);
            Assert.Equal(1, rule.Priority);
            Assert.True(rule.StopIfTrue);
        }

        [Fact]
        public void ExcelRange_ImageExportStopsConditionalRuleDiscoveryAfterBoundedLookahead() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("RuleDiscovery");
            var conditional = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
            };
            conditional.Append(
                new ConditionalFormattingRule { Type = ConditionalFormatValues.Expression, Priority = 2 },
                new ConditionalFormattingRule { Type = ConditionalFormatValues.Expression, Priority = 1 });
            var poison = new ConditionalFormattingRule { Type = ConditionalFormatValues.Expression };
            poison.SetAttribute(new OpenXmlAttribute("priority", string.Empty, "not-an-integer"));
            conditional.Append(poison);
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            worksheet.InsertAfter(conditional, worksheet.GetFirstChild<SheetData>());

            IReadOnlyList<ExcelConditionalFormattingInfo> retained =
                sheet.GetConditionalFormattingRules("A1", 1, out bool truncated);

            Assert.True(truncated);
            Assert.Equal(1, Assert.Single(retained).Priority);
        }

        [Fact]
        public void ExcelRange_ImageExportBoundsSkippedConditionalContainers() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("SkippedContainers");
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            for (int index = 0; index < 10; index++) {
                worksheet.InsertAfter(new ConditionalFormatting {
                    SequenceOfReferences = new ListValue<StringValue> { InnerText = "B" + (index + 1) }
                }, worksheet.GetFirstChild<SheetData>());
            }

            IReadOnlyList<ExcelConditionalFormattingInfo> retained =
                sheet.GetConditionalFormattingRules("A1", 1, out bool truncated);

            Assert.True(truncated);
            Assert.Empty(retained);
        }

        [Fact]
        public void ExcelRange_ImageExportBoundsSkippedConditionalReferenceLists() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("SkippedReferences");
            string references = string.Join(" ", Enumerable.Repeat("B1", 1_000));
            var conditional = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = references }
            };
            conditional.Append(new ConditionalFormattingRule {
                Type = ConditionalFormatValues.Expression,
                Priority = 1
            });
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            worksheet.InsertAfter(conditional, worksheet.GetFirstChild<SheetData>());

            IReadOnlyList<ExcelConditionalFormattingInfo> retained =
                sheet.GetConditionalFormattingRules("A1", 1, out bool truncated);

            Assert.True(truncated);
            Assert.Empty(retained);
        }

        [Fact]
        public void ExcelRange_ImageExportBoundsConditionalRuleCellWork() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("RuleCellWork");
            sheet.CellValue(1, 1, 0D);
            var conditional = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
            };
            for (int priority = 1; priority <= 1_001; priority++) {
                conditional.Append(new ConditionalFormattingRule {
                    Type = ConditionalFormatValues.ColorScale,
                    Priority = priority
                });
            }
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            worksheet.InsertAfter(conditional, worksheet.GetFirstChild<SheetData>());

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A1000").CreateVisualSnapshot();

            OfficeImageExportDiagnostic diagnostic = Assert.Single(
                snapshot.Diagnostics,
                item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalReferenceLimitExceeded);
            Assert.Contains("200-rule", diagnostic.Message, StringComparison.Ordinal);
            Assert.Equal("RuleCellWork!A1:A1000", diagnostic.Source);
        }

        [Fact]
        public void ExcelRange_ImageExportOmitsConditionalRulesWhenNoRuleFitsWorkBudget() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("OversizedRuleCellWork");
            sheet.AddConditionalFormulaRule("A1", "=A1>0", stopIfTrue: true, fillColor: "C6EFCE");
            var cells = new CountingVisualCellList(1_000_001);
            var diagnostics = new List<OfficeImageExportDiagnostic>();

            ExcelConditionalVisualState state = ExcelConditionalVisualEvaluator.Evaluate(
                sheet,
                cells,
                "A1:XFD1048576",
                new DateTime(2026, 1, 1),
                diagnostics);

            Assert.Same(ExcelConditionalVisualState.Empty, state);
            Assert.Equal(0, cells.ReadCount);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal(ExcelImageExportDiagnosticCodes.ConditionalReferenceLimitExceeded, diagnostic.Code);
            Assert.Contains("1000000 rule-cell", diagnostic.Message, StringComparison.Ordinal);
            Assert.Equal("OversizedRuleCellWork!A1:XFD1048576", diagnostic.Source);
        }

        [Fact]
        public void ExcelRange_ImageExportSkipsConditionalRuleDiscoveryWhenNoRuleFitsWorkBudget() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("SkippedRuleDiscovery");
            var conditional = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = "A1" }
            };
            var malformedRule = new ConditionalFormattingRule {
                Type = ConditionalFormatValues.Expression
            };
            malformedRule.SetAttribute(new OpenXmlAttribute("priority", null, "not-an-integer"));
            conditional.Append(malformedRule);
            Worksheet worksheet = sheet.WorksheetPart.Worksheet!;
            worksheet.InsertAfter(conditional, worksheet.GetFirstChild<SheetData>());
            var cells = new CountingVisualCellList(200_001);
            var diagnostics = new List<OfficeImageExportDiagnostic>();

            ExcelConditionalVisualState state = ExcelConditionalVisualEvaluator.Evaluate(
                sheet,
                cells,
                "A1:XFD1048576",
                new DateTime(2026, 1, 1),
                diagnostics);

            Assert.Same(ExcelConditionalVisualState.Empty, state);
            Assert.Equal(0, cells.ReadCount);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal(ExcelImageExportDiagnosticCodes.ConditionalReferenceLimitExceeded, diagnostic.Code);
            Assert.Equal("SkippedRuleDiscovery!A1:XFD1048576", diagnostic.Source);
        }

        [Fact]
        public void ExcelRange_ImageExportReservesEveryConditionalEvaluatorPass() {
            using ExcelDocument document = ExcelDocument.Create(new MemoryStream());
            ExcelSheet sheet = document.AddWorksheet("MultiPassRuleCellWork");
            sheet.AddConditionalFormulaRule("A1", "=A1>0", stopIfTrue: true, fillColor: "C6EFCE");
            var cells = new CountingVisualCellList(1_000_000);
            var diagnostics = new List<OfficeImageExportDiagnostic>();

            ExcelConditionalVisualState state = ExcelConditionalVisualEvaluator.Evaluate(
                sheet,
                cells,
                "A1:XFD1048576",
                new DateTime(2026, 1, 1),
                diagnostics);

            Assert.Same(ExcelConditionalVisualState.Empty, state);
            Assert.Equal(0, cells.ReadCount);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
            Assert.Equal(ExcelImageExportDiagnosticCodes.ConditionalReferenceLimitExceeded, diagnostic.Code);
            Assert.Contains("all evaluator passes", diagnostic.Message, StringComparison.Ordinal);
            Assert.Equal("MultiPassRuleCellWork!A1:XFD1048576", diagnostic.Source);
        }

        [Fact]
        public void ExcelRange_ImageExportKeepsStoppedCellsInColorScaleThresholds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("StoppedColorScale");
            sheet.CellValue(1, 1, 0);
            sheet.CellValue(2, 1, 5);
            sheet.CellValue(3, 1, 10);
            sheet.AddConditionalFormulaRule("A1:A3", "A1=0", stopIfTrue: true);
            sheet.AddConditionalColorScale("A1:A3", OfficeColor.Red, OfficeColor.Lime);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A3").CreateVisualSnapshot();

            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FF808000", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FF00FF00", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 1).Style.FillColorArgb);
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsThreeColorScaleMiddleStop() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("ThreeColor");
                sheet.CellValue(1, 1, 0);
                sheet.CellValue(2, 1, 50);
                sheet.CellValue(3, 1, 100);
                sheet.SetColumnWidth(1, 14);
                sheet.AddConditionalColorScale("A1:A3", OfficeColor.Red, OfficeColor.Green);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                ColorScale colorScale = worksheet.Elements<ConditionalFormatting>().First().Elements<ConditionalFormattingRule>().First().GetFirstChild<ColorScale>()!;
                colorScale.RemoveAllChildren<ConditionalFormatValueObject>();
                colorScale.RemoveAllChildren<X.Color>();
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Number, Val = "0" });
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Number, Val = "50" });
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Number, Val = "100" });
                colorScale.Append(new X.Color { Rgb = "FFFF0000" });
                colorScale.Append(new X.Color { Rgb = "FFFFFF00" });
                colorScale.Append(new X.Color { Rgb = "FF00FF00" });
                worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A3").CreateVisualSnapshot();

                Assert.Equal("FFFF0000", snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.FillColorArgb);
                Assert.Equal("FFFFFF00", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1).Style.FillColorArgb);
                Assert.Equal("FF00FF00", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 1).Style.FillColorArgb);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsPercentileColorScaleMiddleStop() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("PercentileColor");
                int[] values = { 0, 10, 15, 20, 100 };
                for (int row = 1; row <= values.Length; row++) {
                    sheet.CellValue(row, 1, values[row - 1]);
                }

                sheet.AddConditionalColorScale("A1:A5", OfficeColor.Red, OfficeColor.Green);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                ColorScale colorScale = worksheet.Elements<ConditionalFormatting>().First().Elements<ConditionalFormattingRule>().First().GetFirstChild<ColorScale>()!;
                colorScale.RemoveAllChildren<ConditionalFormatValueObject>();
                colorScale.RemoveAllChildren<X.Color>();
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Min });
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Percentile, Val = "50" });
                colorScale.Append(new ConditionalFormatValueObject { Type = ConditionalFormatValueObjectValues.Max });
                colorScale.Append(new X.Color { Rgb = "FFFF0000" });
                colorScale.Append(new X.Color { Rgb = "FFFFFF00" });
                colorScale.Append(new X.Color { Rgb = "FF00FF00" });
                worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelRangeVisualSnapshot snapshot = document.Sheets.Single().Range("A1:A5").CreateVisualSnapshot();

                Assert.Equal("FFFFFF00", snapshot.Cells.Single(cell => cell.Row == 3).Style.FillColorArgb);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRejectsMalformedColorScaleStopCounts() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("MalformedColor");
                for (int row = 1; row <= 4; row++) {
                    sheet.CellValue(row, 1, row);
                }

                sheet.AddConditionalColorScale("A1:A4", OfficeColor.Red, OfficeColor.Green);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                ColorScale colorScale = worksheet.Descendants<ColorScale>().Single();
                colorScale.RemoveAllChildren<ConditionalFormatValueObject>();
                colorScale.RemoveAllChildren<X.Color>();
                for (int index = 0; index < 4; index++) {
                    colorScale.Append(new ConditionalFormatValueObject {
                        Type = ConditionalFormatValueObjectValues.Number,
                        Val = index.ToString(CultureInfo.InvariantCulture)
                    });
                    colorScale.Append(new X.Color { Rgb = "FF" + index.ToString("X6", CultureInfo.InvariantCulture) });
                }

                worksheet.Save();
            }

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelRangeVisualSnapshot snapshot = loaded.Sheets.Single().Range("A1:A4").CreateVisualSnapshot();

            Assert.All(snapshot.Cells, cell => Assert.Null(cell.Style.FillColorArgb));
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsConditionalIconSetHiddenValues() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("IconOnly");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 2);
            sheet.CellValue(3, 1, 3);
            sheet.SetColumnWidth(1, 12);
            sheet.SetRowHeight(1, 24);
            sheet.SetRowHeight(2, 24);
            sheet.SetRowHeight(3, 24);
            sheet.AddConditionalIconSet("A1:A3", IconSetValues.ThreeTrafficLights1, showValue: false, reverseIconOrder: false);

            ExcelRange range = sheet.Range("A1:A3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(3, snapshot.ConditionalIcons.Count);
            Assert.All(snapshot.ConditionalIcons, icon => Assert.False(icon.ShowValue));
            Assert.DoesNotContain(">1<", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(">2<", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(">3<", svg, StringComparison.Ordinal);
            Assert.Contains("#16A34A", svg, StringComparison.Ordinal);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, diagnostic.Severity);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            ExcelVisualConditionalIcon finalIcon = snapshot.ConditionalIcons.Single(icon => icon.Row == 3 && icon.Column == 1);
            Assert.True(CountGreenIconPixels(rendered!, finalIcon) > 4, "Expected visible green conditional-formatting icon pixels.");
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsConditionalIconSetStrictThresholds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("IconThresholds");
                sheet.CellValue(1, 1, 0);
                sheet.CellValue(2, 1, 50);
                sheet.CellValue(3, 1, 100);
                sheet.AddConditionalIconSet("A1:A3", IconSetValues.ThreeTrafficLights1, showValue: true, reverseIconOrder: false);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                IconSet iconSet = worksheet.Elements<ConditionalFormatting>().First().Elements<ConditionalFormattingRule>().First().GetFirstChild<IconSet>()!;
                ConditionalFormatValueObject[] thresholds = iconSet.Elements<ConditionalFormatValueObject>().ToArray();
                thresholds[1].Type = ConditionalFormatValueObjectValues.Number;
                thresholds[1].Val = "50";
                thresholds[1].GreaterThanOrEqual = false;
                thresholds[2].Type = ConditionalFormatValueObjectValues.Number;
                thresholds[2].Val = "100";
                thresholds[2].GreaterThanOrEqual = true;
                worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A3").CreateVisualSnapshot();

                ExcelConditionalFormattingInfo info = Assert.Single(sheet.GetConditionalFormattingRules("A1:A3"));
                Assert.False(info.IconSetThresholds[1].GreaterThanOrEqual);
                Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 1 && icon.Kind == ExcelConditionalIconKind.RedCircle);
                Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 2 && icon.Kind == ExcelConditionalIconKind.RedCircle);
                Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 3 && icon.Kind == ExcelConditionalIconKind.GreenCircle);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsConditionalIconSetPercentileThresholds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("IconPercentiles");
                sheet.CellValue(1, 1, 1);
                sheet.CellValue(2, 1, 2);
                sheet.CellValue(3, 1, 3);
                sheet.CellValue(4, 1, 100);
                sheet.AddConditionalIconSet("A1:A4", IconSetValues.ThreeTrafficLights1, showValue: true, reverseIconOrder: false);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                IconSet iconSet = worksheet.Elements<ConditionalFormatting>().First().Elements<ConditionalFormattingRule>().First().GetFirstChild<IconSet>()!;
                ConditionalFormatValueObject[] thresholds = iconSet.Elements<ConditionalFormatValueObject>().ToArray();
                thresholds[1].Type = ConditionalFormatValueObjectValues.Percentile;
                thresholds[1].Val = "50";
                thresholds[2].Type = ConditionalFormatValueObjectValues.Percentile;
                thresholds[2].Val = "90";
                worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A4").CreateVisualSnapshot();

                ExcelConditionalFormattingInfo info = Assert.Single(sheet.GetConditionalFormattingRules("A1:A4"));
                Assert.Equal("percentile", info.IconSetThresholds[1].Type, ignoreCase: true);
                Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 2 && icon.Kind == ExcelConditionalIconKind.RedCircle);
                Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 3 && icon.Kind == ExcelConditionalIconKind.YellowCircle);
                Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 4 && icon.Kind == ExcelConditionalIconKind.GreenCircle);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRejectsNonFiniteIconSetPercentiles() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            try {
                using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                    ExcelSheet sheet = document.AddWorksheet("IconPercentiles");
                    sheet.CellValue(1, 1, 1);
                    sheet.CellValue(2, 1, 2);
                    sheet.CellValue(3, 1, 3);
                    sheet.AddConditionalIconSet("A1:A3", IconSetValues.ThreeTrafficLights1,
                        showValue: true, reverseIconOrder: false);
                    document.Save();
                }

                using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                    Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                    IconSet iconSet = worksheet.Elements<ConditionalFormatting>()
                        .First().Elements<ConditionalFormattingRule>().First().GetFirstChild<IconSet>()!;
                    ConditionalFormatValueObject threshold = iconSet.Elements<ConditionalFormatValueObject>().ElementAt(1);
                    threshold.Type = ConditionalFormatValueObjectValues.Percentile;
                    threshold.Val = "NaN";
                    worksheet.Save();
                }

                using ExcelDocument reopened = ExcelDocument.Load(filePath);
                ExcelRangeVisualSnapshot snapshot = reopened.Sheets.Single().Range("A1:A3").CreateVisualSnapshot();

                Assert.Equal(3, snapshot.ConditionalIcons.Count);
            } finally {
                if (File.Exists(filePath)) File.Delete(filePath);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsConditionalDataBarHiddenValues() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("DataBarOnly");
                sheet.CellValue(1, 1, 42);
                sheet.SetColumnWidth(1, 12);
                sheet.SetRowHeight(1, 24);
                sheet.AddConditionalDataBar("A1:A1", OfficeColor.Blue);
                document.Save();
            }

            using (SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true)) {
                Worksheet worksheet = spreadsheet.WorkbookPart!.WorksheetParts.First().Worksheet;
                DataBar dataBar = worksheet.Elements<ConditionalFormatting>().First().Elements<ConditionalFormattingRule>().First().GetFirstChild<DataBar>()!;
                dataBar.ShowValue = false;
                worksheet.Save();
            }

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRange range = sheet.Range("A1:A1");
                ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
                string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

                ExcelVisualConditionalDataBar bar = Assert.Single(snapshot.ConditionalDataBars);
                Assert.False(bar.ShowValue);
                Assert.DoesNotContain(">42<", svg, StringComparison.Ordinal);
                Assert.Contains("#0000FF", svg, StringComparison.Ordinal);
            }
        }

        [Fact]
        public void ExcelRange_ImageExportRendersFirstPriorityDataBarForOverlappingRules() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("DataBarPriority");
            sheet.CellValue(1, 1, 20);
            sheet.CellValue(2, 1, 10);
            sheet.SetColumnWidth(1, 12);
            sheet.AddConditionalDataBar("A1:A2", OfficeColor.Blue);
            sheet.AddConditionalDataBar("A1:A1", OfficeColor.Red);

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(new[] { "FF0000FF", "FFFF0000" }, snapshot.ConditionalDataBars.Select(bar => bar.ColorArgb).ToArray());
            Assert.Contains("#0000FF", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("#FF0000", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportUsesRawNumericValuesForFormattedConditionalCandidates() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Formatted");
            for (int row = 1; row <= 3; row++) {
                double value = row / 10D;
                sheet.CellValue(row, 1, value);
                sheet.CellValue(row, 2, value);
                ApplyBuiltInNumberFormatId(document, sheet, "A" + row.ToString(CultureInfo.InvariantCulture), 10U);
                ApplyBuiltInNumberFormatId(document, sheet, "B" + row.ToString(CultureInfo.InvariantCulture), 10U);
            }

            sheet.AddConditionalDataBar("A1:A3", OfficeColor.Blue);
            sheet.AddConditionalIconSet("B1:B3", IconSetValues.ThreeTrafficLights1, showValue: true, reverseIconOrder: false);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:B3").CreateVisualSnapshot();

            Assert.Equal(new[] { "10.00%", "20.00%", "30.00%" }, snapshot.Cells.Where(cell => cell.Column == 1).Select(cell => cell.Text).ToArray());
            Assert.Equal(3, snapshot.ConditionalDataBars.Count);
            Assert.Equal(3, snapshot.ConditionalIcons.Count);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 3 && icon.Column == 2 && icon.Kind == ExcelConditionalIconKind.GreenCircle);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesAbsoluteConditionalFormulaReferences() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Absolute");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 0);
            sheet.CellValue(3, 1, 0);
            for (int row = 1; row <= 3; row++) {
                sheet.CellValue(row, 2, "row " + row.ToString(CultureInfo.InvariantCulture));
            }

            sheet.AddConditionalFormulaRule("B1:B3", "=$A$1>0", stopIfTrue: false, fillColor: "C6EFCE");

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:B3").CreateVisualSnapshot();

            Assert.All(snapshot.Cells.Where(cell => cell.Column == 2), cell => Assert.Equal("FFC6EFCE", cell.Style.FillColorArgb));
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsStopIfTrueRulesWithoutSupportedFills() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("StopOnly");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(1, 2, 10);
            sheet.AddConditionalFormulaRule("B1:B1", "=$A$1>0", stopIfTrue: true, fillColor: null);
            sheet.AddConditionalFormulaRule("B1:B1", "=B1>0", stopIfTrue: false, fillColor: "FCE4D6");

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:B1").CreateVisualSnapshot();

            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 2).Style.FillColorArgb);
        }

        [Fact]
        public void ExcelRange_ImageExportStopIfTrueSuppressesLowerIconsAndDataBars() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("StopVisuals");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 2);
            sheet.CellValue(3, 1, 3);
            sheet.AddConditionalFormulaRule("A2:A2", "=A2>0", stopIfTrue: true, fillColor: null);
            sheet.AddConditionalDataBar("A1:A3", OfficeColor.Blue);
            sheet.AddConditionalIconSet("A1:A3", IconSetValues.ThreeTrafficLights1, showValue: true, reverseIconOrder: false);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A3").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(2, snapshot.ConditionalDataBars.Count);
            Assert.DoesNotContain(snapshot.ConditionalDataBars, bar => bar.Row == 2);
            Assert.Equal(2, snapshot.ConditionalIcons.Count);
            Assert.DoesNotContain(snapshot.ConditionalIcons, icon => icon.Row == 2);
        }

        [Fact]
        public void ExcelRange_ImageExportRendersFiveIconConditionalSetsWithApproximationDiagnostic() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("FiveIcons");
            for (int row = 1; row <= 5; row++) {
                sheet.CellValue(row, 1, row);
                sheet.SetRowHeight(row, 24);
            }

            sheet.SetColumnWidth(1, 12);
            sheet.AddConditionalIconSet("A1:A5", IconSetValues.FiveArrows, showValue: true, reverseIconOrder: false);

            ExcelRange range = sheet.Range("A1:A5");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(5, snapshot.ConditionalIcons.Count);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 1 && icon.Kind == ExcelConditionalIconKind.RedDownArrow);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 2 && icon.Kind == ExcelConditionalIconKind.YellowDownArrow);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 3 && icon.Kind == ExcelConditionalIconKind.YellowSideArrow);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 4 && icon.Kind == ExcelConditionalIconKind.YellowUpArrow);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 5 && icon.Kind == ExcelConditionalIconKind.GreenUpArrow);
            Assert.Contains("#16A34A", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalIconSetUnsupported);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, diagnostic.Severity);
            Assert.Equal("FiveIcons!A1:A5", diagnostic.Source);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            ExcelVisualConditionalIcon finalIcon = snapshot.ConditionalIcons.Single(icon => icon.Row == 5 && icon.Column == 1);
            Assert.True(CountGreenIconPixels(rendered!, finalIcon) > 4, "Expected visible green conditional-formatting arrow pixels.");
        }

        [Fact]
        public void ExcelRange_ImageExportRendersFlagIconConditionalSetsWithSharedDrawingIcons() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("FlagIcons");
            for (int row = 1; row <= 3; row++) {
                sheet.CellValue(row, 1, row);
                sheet.SetRowHeight(row, 24);
            }

            sheet.SetColumnWidth(1, 12);
            sheet.AddConditionalIconSet("A1:A3", IconSetValues.ThreeFlags, showValue: true, reverseIconOrder: false);

            ExcelRange range = sheet.Range("A1:A3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(3, snapshot.ConditionalIcons.Count);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 1 && icon.Kind == ExcelConditionalIconKind.RedFlag);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 2 && icon.Kind == ExcelConditionalIconKind.YellowFlag);
            Assert.Contains(snapshot.ConditionalIcons, icon => icon.Row == 3 && icon.Kind == ExcelConditionalIconKind.GreenFlag);
            Assert.Contains("#16A34A", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalIconSetUnsupported);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalIconSetApproximation);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, diagnostic.Severity);
            Assert.Equal("FlagIcons!A1:A3", diagnostic.Source);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            ExcelVisualConditionalIcon finalIcon = snapshot.ConditionalIcons.Single(icon => icon.Row == 3 && icon.Column == 1);
            Assert.True(CountGreenIconPixels(rendered!, finalIcon) > 4, "Expected visible green conditional-formatting flag pixels.");
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesConditionalCellIsAndFormulaDifferentialFills() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Rules");
            sheet.CellValue(1, 1, 5);
            sheet.CellValue(2, 1, 20);
            sheet.CellValue(3, 1, 30);
            sheet.CellValue(1, 2, -5);
            sheet.CellValue(2, 2, 0);
            sheet.CellValue(3, 2, 5);
            sheet.SetColumnWidth(1, 12);
            sheet.SetColumnWidth(2, 12);
            sheet.SetRowHeight(1, 24);
            sheet.SetRowHeight(2, 24);
            sheet.SetRowHeight(3, 24);
            sheet.AddConditionalRule("A1:A3", ConditionalFormattingOperatorValues.GreaterThan, "10", fillColor: "C6EFCE");
            sheet.AddConditionalFormulaRule("B1:B3", "B1<0", stopIfTrue: true, fillColor: "FEE2E2");
            sheet.AddConditionalColorScale("B1:B3", OfficeColor.Red, OfficeColor.Lime);

            ExcelRange range = sheet.Range("A1:B3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            IReadOnlyList<ExcelConditionalFormattingInfo> rules = sheet.GetConditionalFormattingRules("A1:B3");
            ExcelConditionalFormattingInfo cellIsRule = Assert.Single(rules, rule => rule.Type == "CellIs");
            ExcelConditionalFormattingInfo formulaRule = Assert.Single(rules, rule => rule.Type == "Expression");
            Assert.NotNull(cellIsRule.DifferentialFormatId);
            Assert.Equal("FFC6EFCE", cellIsRule.DifferentialFillColorArgb);
            Assert.True(formulaRule.StopIfTrue);
            Assert.Equal("FFFEE2E2", formulaRule.DifferentialFillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFC6EFCE", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFC6EFCE", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFFEE2E2", snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FF808000", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FF00FF00", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 2).Style.FillColorArgb);
            Assert.Contains("#C6EFCE", svg, StringComparison.Ordinal);
            Assert.Contains("#FEE2E2", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            ExcelVisualCell cellIsRendered = snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1);
            ExcelVisualCell formulaRendered = snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 2);
            AssertCellContainsPixelNear(
                rendered!,
                cellIsRendered,
                OfficeColor.FromRgb(198, 239, 206),
                tolerance: 3);
            AssertCellContainsPixelNear(
                rendered!,
                formulaRendered,
                OfficeColor.FromRgb(254, 226, 226),
                tolerance: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesConditionalDifferentialFontStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("FontRules");
            sheet.CellValue(1, 1, 5);
            sheet.CellValue(2, 1, 20);
            sheet.CellValue(3, 1, 30);
            sheet.SetColumnWidth(1, 16);
            sheet.SetRowHeight(1, 24);
            sheet.SetRowHeight(2, 24);
            sheet.SetRowHeight(3, 24);
            sheet.AddConditionalRule("A1:A3", ConditionalFormattingOperatorValues.GreaterThan, "10");
            uint differentialFormatId = sheet.AppendConditionalDifferentialFormat(new X.DifferentialFormat(
                new X.Font(
                    new X.Bold(),
                    new X.Italic(),
                    new X.Underline(),
                    new X.FontSize { Val = 18D },
                    new X.FontName { Val = "Aptos Display" },
                    new X.Color { Rgb = "FFFF0000" })));
            sheet.SetLastConditionalFormattingRuleDifferentialFormatId(differentialFormatId);

            ExcelRange range = sheet.Range("A1:A3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            ExcelConditionalFormattingInfo rule = Assert.Single(sheet.GetConditionalFormattingRules("A1:A3"));
            Assert.Equal("FFFF0000", rule.DifferentialFontColorArgb);
            Assert.True(rule.DifferentialFontBold);
            Assert.True(rule.DifferentialFontItalic);
            Assert.True(rule.DifferentialFontUnderline);
            Assert.Equal("Aptos Display", rule.DifferentialFontName);
            Assert.Equal(18D, rule.DifferentialFontSize);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.FontColorArgb);
            ExcelVisualCell firstMatched = snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1);
            Assert.Equal("FFFF0000", firstMatched.Style.FontColorArgb);
            Assert.True(firstMatched.Style.Bold);
            Assert.True(firstMatched.Style.Italic);
            Assert.True(firstMatched.Style.Underline);
            Assert.Equal("Aptos Display", firstMatched.Style.FontName);
            Assert.Equal(18D, firstMatched.Style.FontSize);
            Assert.Equal("FFFF0000", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 1).Style.FontColorArgb);
            Assert.Contains("fill=\"#FF0000\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-family=\"Aptos Display", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-size=\"18\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("text-decoration=\"underline\"", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalDifferentialFormatUnsupported);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesConditionalDifferentialBorders() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("BorderRules");
            sheet.CellValue(1, 1, 5);
            sheet.CellValue(2, 1, 20);
            sheet.SetColumnWidth(1, 16);
            sheet.SetRowHeight(1, 28);
            sheet.SetRowHeight(2, 28);
            sheet.AddConditionalRule("A1:A2", ConditionalFormattingOperatorValues.GreaterThan, "10");
            uint differentialFormatId = sheet.AppendConditionalDifferentialFormat(new X.DifferentialFormat(
                new X.Border(
                    new X.LeftBorder(new X.Color { Rgb = "FFFF0000" }) { Style = X.BorderStyleValues.Thick },
                    new X.RightBorder(new X.Color { Rgb = "FFFF0000" }) { Style = X.BorderStyleValues.Thick },
                    new X.TopBorder(new X.Color { Rgb = "FFFF0000" }) { Style = X.BorderStyleValues.Thick },
                    new X.BottomBorder(new X.Color { Rgb = "FFFF0000" }) { Style = X.BorderStyleValues.Thick })));
            sheet.SetLastConditionalFormattingRuleDifferentialFormatId(differentialFormatId);

            var options = new ExcelImageExportOptions { ShowGridlines = false };
            ExcelRange range = sheet.Range("A1:A2");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            ExcelConditionalFormattingInfo rule = Assert.Single(sheet.GetConditionalFormattingRules("A1:A2"));
            Assert.NotNull(rule.DifferentialBorder);
            Assert.Equal("thick", rule.DifferentialBorder!.Top!.Style);
            Assert.Equal("FFFF0000", rule.DifferentialBorder.Top.ColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.Border);
            ExcelVisualCell matched = snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1);
            Assert.Equal("thick", matched.Style.Border!.Top!.Style);
            Assert.Equal("FFFF0000", matched.Style.Border.Top.ColorArgb);
            Assert.Contains("#FF0000", svg, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalDifferentialFormatUnsupported);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    matched.X + 2D,
                    matched.Y,
                    matched.X + matched.Width - 2D,
                    matched.Y + 5D,
                    OfficeColor.Red,
                    tolerance: 3),
                "Expected the conditional top border to render as red pixels in the PNG artifact.");
        }

        [Fact]
        public void ExcelRange_ImageExportRespectsFalseConditionalDifferentialFontFlags() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("FalseFontRules");
            sheet.CellValue(1, 1, 20);
            sheet.AddConditionalRule("A1:A1", ConditionalFormattingOperatorValues.GreaterThan, "10");

            var bold = new X.Bold();
            bold.SetAttribute(new OpenXmlAttribute("val", string.Empty, "0"));
            var italic = new X.Italic();
            italic.SetAttribute(new OpenXmlAttribute("val", string.Empty, "0"));
            var underline = new X.Underline();
            underline.SetAttribute(new OpenXmlAttribute("val", string.Empty, "0"));
            uint differentialFormatId = sheet.AppendConditionalDifferentialFormat(new X.DifferentialFormat(
                new X.Font(bold, italic, underline, new X.Color { Rgb = "FFFF0000" })));
            sheet.SetLastConditionalFormattingRuleDifferentialFormatId(differentialFormatId);

            ExcelConditionalFormattingInfo rule = Assert.Single(sheet.GetConditionalFormattingRules("A1:A1"));
            ExcelVisualCell cell = Assert.Single(sheet.Range("A1:A1").CreateVisualSnapshot().Cells);

            Assert.False(rule.DifferentialFontBold);
            Assert.False(rule.DifferentialFontItalic);
            Assert.False(rule.DifferentialFontUnderline);
            Assert.False(cell.Style.Bold);
            Assert.False(cell.Style.Italic);
            Assert.False(cell.Style.Underline);
            Assert.Equal("FFFF0000", cell.Style.FontColorArgb);
        }

        [Fact]
        public void ExcelRange_ImageExportTreatsConditionalDifferentialUnderlineNoneAsDisabled() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("UnderlineNone");
            sheet.CellValue(1, 1, 20);
            sheet.AddConditionalRule("A1:A1", ConditionalFormattingOperatorValues.GreaterThan, "10");

            uint differentialFormatId = sheet.AppendConditionalDifferentialFormat(new X.DifferentialFormat(
                new X.Font(new X.Underline { Val = X.UnderlineValues.None })));
            sheet.SetLastConditionalFormattingRuleDifferentialFormatId(differentialFormatId);

            ExcelConditionalFormattingInfo rule = Assert.Single(sheet.GetConditionalFormattingRules("A1:A1"));
            ExcelVisualCell cell = Assert.Single(sheet.Range("A1:A1").CreateVisualSnapshot().Cells);

            Assert.False(rule.DifferentialFontUnderline);
            Assert.False(cell.Style.Underline);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnsupportedConditionalRuleShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Unsupported");
            sheet.CellValue(1, 1, "Hot");
            sheet.CellValue(2, 1, "Warm");
            sheet.CellValue(1, 2, 2);
            sheet.CellValue(2, 2, 4);
            sheet.SetColumnWidth(1, 12);
            sheet.SetColumnWidth(2, 12);
            sheet.AddConditionalRule("A1:A2", ConditionalFormattingOperatorValues.Equal, "\"Hot\"", fillColor: "FEE2E2");
            sheet.AddConditionalFormulaRule("B1:B2", "MOD(B1,2)=0", fillColor: "C6EFCE");
            AddMalformedTimePeriodRuleWithFill(sheet, "B1:B2");

            ExcelRange range = sheet.Range("A1:B2");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 2).Style.FillColorArgb);
            OfficeImageExportDiagnostic cellIs = Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalCellIsUnsupported);
            OfficeImageExportDiagnostic formula = Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalFormulaUnsupported);
            OfficeImageExportDiagnostic timePeriod = Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalTimePeriodUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, cellIs.Severity);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, formula.Severity);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, timePeriod.Severity);
            Assert.Equal("Unsupported!A1:A2", cellIs.Source);
            Assert.Equal("Unsupported!B1:B2", formula.Source);
            Assert.Equal("Unsupported!B1:B2", timePeriod.Source);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesTextConditionalFills() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("TextRules");
            string[] containsValues = { "Hot path", "Warm path", "cold path", "preheat" };
            string[] beginsValues = { "Ops North", "Dev North", "Ops South", "ops west" };
            string[] endsValues = { "North Ops", "Ops North", "West ops", "Ops" };
            for (int row = 1; row <= containsValues.Length; row++) {
                sheet.CellValue(row, 1, containsValues[row - 1]);
                sheet.CellValue(row, 2, containsValues[row - 1]);
                sheet.CellValue(row, 3, beginsValues[row - 1]);
                sheet.CellValue(row, 4, endsValues[row - 1]);
            }

            for (int column = 1; column <= 4; column++) {
                sheet.SetColumnWidth(column, 15);
            }

            sheet.Range("A1:A4").ConditionalFormatting.ContainsText("hot", "FCE4D6");
            sheet.Range("B1:B4").ConditionalFormatting.NotContainsText("hot", "DBEAFE");
            sheet.Range("C1:C4").ConditionalFormatting.BeginsWithText("ops", "C6EFCE");
            sheet.Range("D1:D4").ConditionalFormatting.EndsWithText("ops", "FEE2E2");

            ExcelRange range = sheet.Range("A1:D4");
            ExcelImageExportOptions options = new ExcelImageExportOptions { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("FFFCE4D6", snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FFC6EFCE", snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 3).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 3).Style.FillColorArgb);
            Assert.Equal("FFC6EFCE", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 3).Style.FillColorArgb);
            Assert.Equal("FFC6EFCE", snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 3).Style.FillColorArgb);
            Assert.Equal("FFFEE2E2", snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 4).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 4).Style.FillColorArgb);
            Assert.Equal("FFFEE2E2", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 4).Style.FillColorArgb);
            Assert.Equal("FFFEE2E2", snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 4).Style.FillColorArgb);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalTextRuleUnsupported);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalRuleUnsupported);
            Assert.Contains("#FCE4D6", svg, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svg, StringComparison.Ordinal);
            Assert.Contains("#C6EFCE", svg, StringComparison.Ordinal);
            Assert.Contains("#FEE2E2", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);

            ExcelVisualCell containsCell = snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1);
            ExcelVisualCell notContainsCell = snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 2);
            ExcelVisualCell beginsCell = snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 3);
            ExcelVisualCell endsCell = snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 4);
            AssertPixelNear(
                rendered!,
                (int)(containsCell.X + containsCell.Width - 8),
                (int)(containsCell.Y + containsCell.Height - 8),
                OfficeColor.FromRgb(252, 228, 214),
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(notContainsCell.X + notContainsCell.Width - 8),
                (int)(notContainsCell.Y + notContainsCell.Height - 8),
                OfficeColor.FromRgb(219, 234, 254),
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(beginsCell.X + beginsCell.Width - 8),
                (int)(beginsCell.Y + beginsCell.Height - 8),
                OfficeColor.FromRgb(198, 239, 206),
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(endsCell.X + endsCell.Width - 8),
                (int)(endsCell.Y + endsCell.Height - 8),
                OfficeColor.FromRgb(254, 226, 226),
                tolerance: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesAboveBelowAverageConditionalFillsAndDiagnosesStdDevVariant() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Averages");
            int[] values = { 1, 2, 3, 4, 5 };
            for (int row = 1; row <= values.Length; row++) {
                sheet.CellValue(row, 1, values[row - 1]);
                sheet.CellValue(row, 2, values[row - 1]);
                sheet.CellValue(row, 3, values[row - 1]);
            }

            sheet.SetColumnWidth(1, 12);
            sheet.SetColumnWidth(2, 12);
            sheet.SetColumnWidth(3, 12);
            sheet.Range("A1:A5").ConditionalFormatting.AboveAverage("C6EFCE");
            sheet.Range("B1:B5").ConditionalFormatting.BelowAverage("DBEAFE", equalAverage: true);
            sheet.AddConditionalAboveAverageRule("C1:C5", aboveAverage: true, equalAverage: false, fillColor: "FCE4D6");
            MarkAboveAverageRuleAsStdDev(sheet, "C1:C5", stdDev: 1);

            ExcelRange range = sheet.Range("A1:C5");
            ExcelImageExportOptions options = new ExcelImageExportOptions { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFC6EFCE", snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFC6EFCE", snapshot.Cells.Single(cell => cell.Row == 5 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 2).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 2).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 5 && cell.Column == 3).Style.FillColorArgb);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalAboveAverageUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Averages!C1:C5", diagnostic.Source);
            Assert.Contains("#C6EFCE", svg, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("#FCE4D6", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            ExcelVisualCell aboveCell = snapshot.Cells.Single(cell => cell.Row == 5 && cell.Column == 1);
            ExcelVisualCell belowCell = snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 2);
            AssertCellContainsPixelNear(
                rendered!,
                aboveCell,
                OfficeColor.FromRgb(198, 239, 206),
                tolerance: 3);
            AssertCellContainsPixelNear(
                rendered!,
                belowCell,
                OfficeColor.FromRgb(219, 234, 254),
                tolerance: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportCalculatesAboveAverageBeforeStopIfTrueSuppression() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("StoppedAverage");
            sheet.CellValue(1, 1, 1);
            sheet.CellValue(2, 1, 2);
            sheet.CellValue(3, 1, 100);
            sheet.AddConditionalFormulaRule("A3:A3", "=A3>0", stopIfTrue: true, fillColor: "FCE4D6");
            sheet.Range("A1:A3").ConditionalFormatting.AboveAverage("C6EFCE");

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A3").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 2).Style.FillColorArgb);
            Assert.Equal("FFFCE4D6", snapshot.Cells.Single(cell => cell.Row == 3).Style.FillColorArgb);
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesDuplicateAndUniqueValueConditionalFills() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Distinct");
            string[] values = { "Alpha", "Beta", "alpha", string.Empty, "Gamma", "Beta" };
            for (int row = 1; row <= values.Length; row++) {
                sheet.CellValue(row, 1, values[row - 1]);
                sheet.CellValue(row, 2, values[row - 1]);
            }

            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 14);
            sheet.Range("A1:A6").ConditionalFormatting.DuplicateValues("FCE4D6");
            sheet.Range("B1:B6").ConditionalFormatting.UniqueValues("DBEAFE");

            ExcelRange range = sheet.Range("A1:B6");
            ExcelImageExportOptions options = new ExcelImageExportOptions { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("FFFCE4D6", snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFFCE4D6", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFFCE4D6", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 5 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFFCE4D6", snapshot.Cells.Single(cell => cell.Row == 6 && cell.Column == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 2).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 2).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 2).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 5 && cell.Column == 2).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 6 && cell.Column == 2).Style.FillColorArgb);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalRuleUnsupported);
            Assert.Contains("#FCE4D6", svg, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            ExcelVisualCell alphaCell = snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1);
            ExcelVisualCell betaCell = snapshot.Cells.Single(cell => cell.Row == 6 && cell.Column == 1);
            ExcelVisualCell gammaCell = snapshot.Cells.Single(cell => cell.Row == 5 && cell.Column == 2);
            AssertPixelNear(
                rendered!,
                (int)(alphaCell.X + alphaCell.Width - 8),
                (int)(alphaCell.Y + alphaCell.Height - 8),
                OfficeColor.FromRgb(252, 228, 214),
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(betaCell.X + betaCell.Width - 8),
                (int)(betaCell.Y + betaCell.Height - 8),
                OfficeColor.FromRgb(252, 228, 214),
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(gammaCell.X + gammaCell.Width - 8),
                (int)(gammaCell.Y + gammaCell.Height - 8),
                OfficeColor.FromRgb(219, 234, 254),
                tolerance: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportEvaluatesDuplicateRulesAgainstFullRuleRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("PartialDistinct");
            sheet.CellValue(1, 1, "Alpha");
            sheet.CellValue(2, 1, "Alpha");
            sheet.CellValue(3, 1, "Beta");
            sheet.Range("A1:A3").ConditionalFormatting.DuplicateValues("FCE4D6");

            ExcelRangeVisualSnapshot duplicateSnapshot = sheet.Range("A1:A1").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelRangeVisualSnapshot uniqueSnapshot = sheet.Range("A3:A3").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal("FFFCE4D6", Assert.Single(duplicateSnapshot.Cells).Style.FillColorArgb);
            Assert.Null(Assert.Single(uniqueSnapshot.Cells).Style.FillColorArgb);
        }

        [Fact]
        public void ExcelRange_ImageExportComparesDuplicateRulesUsingRawCellValues() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("RawDistinct");
            sheet.Cell(1, 1, 1, numberFormat: "0.00");
            sheet.Cell(2, 1, 1, numberFormat: "0");
            sheet.Cell(3, 1, 2, numberFormat: "0.00");
            sheet.Range("A1:A3").ConditionalFormatting.DuplicateValues("FCE4D6");

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A3").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal("1.00", snapshot.Cells.Single(cell => cell.Row == 1).Text);
            Assert.Equal("1", snapshot.Cells.Single(cell => cell.Row == 2).Text);
            Assert.Equal("FFFCE4D6", snapshot.Cells.Single(cell => cell.Row == 1).Style.FillColorArgb);
            Assert.Equal("FFFCE4D6", snapshot.Cells.Single(cell => cell.Row == 2).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 3).Style.FillColorArgb);
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesTopBottomConditionalFillsIncludingPercentVariants() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("TopBottom");
            int[] topValues = { 10, 40, 30, 40, 20 };
            int[] bottomValues = { 1, 2, -5, 0, -5 };
            int[] topPercentValues = { 1, 5, 3, 5, 2 };
            int[] bottomPercentValues = { 8, 1, 3, 1, 9 };
            for (int row = 1; row <= 5; row++) {
                sheet.CellValue(row, 1, topValues[row - 1]);
                sheet.CellValue(row, 2, bottomValues[row - 1]);
                sheet.CellValue(row, 3, topPercentValues[row - 1]);
                sheet.CellValue(row, 4, bottomPercentValues[row - 1]);
            }

            sheet.SetColumnWidth(1, 12);
            sheet.SetColumnWidth(2, 12);
            sheet.SetColumnWidth(3, 12);
            sheet.SetColumnWidth(4, 12);
            sheet.Range("A1:A5").ConditionalFormatting.Top(2, "C6EFCE");
            sheet.Range("B1:B5").ConditionalFormatting.Bottom(1, "FEE2E2");
            sheet.AddConditionalTopBottomRule("C1:C5", 20, bottom: false, percent: true, fillColor: "FCE4D6");
            sheet.AddConditionalTopBottomRule("D1:D5", 40, bottom: true, percent: true, fillColor: "DBEAFE");

            ExcelRange range = sheet.Range("A1:D5");
            ExcelImageExportOptions options = new ExcelImageExportOptions { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("FFC6EFCE", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFC6EFCE", snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 1).Style.FillColorArgb);
            Assert.Equal("FFFEE2E2", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FFFEE2E2", snapshot.Cells.Single(cell => cell.Row == 5 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FFFCE4D6", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 3).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 3).Style.FillColorArgb);
            Assert.Equal("FFFCE4D6", snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 3).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 4).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 4).Style.FillColorArgb);
            Assert.Equal("FFDBEAFE", snapshot.Cells.Single(cell => cell.Row == 4 && cell.Column == 4).Style.FillColorArgb);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalTopBottomUnsupported);
            Assert.Contains("#C6EFCE", svg, StringComparison.Ordinal);
            Assert.Contains("#FEE2E2", svg, StringComparison.Ordinal);
            Assert.Contains("#FCE4D6", svg, StringComparison.Ordinal);
            Assert.Contains("#DBEAFE", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            ExcelVisualCell topCell = snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1);
            ExcelVisualCell bottomCell = snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 2);
            ExcelVisualCell topPercentCell = snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 3);
            ExcelVisualCell bottomPercentCell = snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 4);
            AssertCellContainsPixelNear(
                rendered!,
                topCell,
                OfficeColor.FromRgb(198, 239, 206),
                tolerance: 3);
            AssertCellContainsPixelNear(
                rendered!,
                bottomCell,
                OfficeColor.FromRgb(254, 226, 226),
                tolerance: 3);
            AssertCellContainsPixelNear(
                rendered!,
                topPercentCell,
                OfficeColor.FromRgb(252, 228, 214),
                tolerance: 3);
            AssertCellContainsPixelNear(
                rendered!,
                bottomPercentCell,
                OfficeColor.FromRgb(219, 234, 254),
                tolerance: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportRanksTopBottomAgainstFullRuleRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("PartialTop");
            for (int row = 1; row <= 10; row++) {
                sheet.CellValue(row, 1, row);
            }

            sheet.Range("A1:A10").ConditionalFormatting.Top(1, "C6EFCE");

            ExcelRangeVisualSnapshot middleSnapshot = sheet.Range("A5:A5").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelRangeVisualSnapshot topSnapshot = sheet.Range("A10:A10").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Null(Assert.Single(middleSnapshot.Cells).Style.FillColorArgb);
            Assert.Equal("FFC6EFCE", Assert.Single(topSnapshot.Cells).Style.FillColorArgb);
        }

        [Fact]
        public void ExcelSheet_ExportsUsedRangeToPngAndSvg() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Summary");
            sheet.CellValue(1, 1, "Metric");
            sheet.CellValue(1, 2, "Score");
            sheet.CellValue(2, 1, "Quality");
            sheet.CellValue(2, 2, 98);

            byte[] png = sheet.ToPng();
            string svg = sheet.ToSvg();

            OfficeImageInfo info = OfficeImageReader.Identify(png);
            Assert.Equal(OfficeImageFormat.Png, info.Format);
            Assert.True(info.Width > 0);
            Assert.True(info.Height > 0);
            Assert.Contains("Quality", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ExportsEmbeddedPngImages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Images");
            sheet.CellValue(1, 1, "Name");
            sheet.CellValue(2, 1, "Logo");
            byte[] badge = CreateSolidPng(12, 10, OfficeColor.FromRgb(220, 38, 38));
            sheet.AddImage(2, 2, badge, "image/png", widthPixels: 12, heightPixels: 10, name: "Badge");

            ExcelRange range = sheet.Range("A1:C4");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult sheetPng = sheet.ExportImage(OfficeImageExportFormat.Png);
            string svg = range.ToSvg();

            Assert.Single(snapshot.Images);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            OfficeColor imagePixel = rendered!.GetPixel(65, 21);
            Assert.True(imagePixel.R > 180);
            Assert.True(imagePixel.G < 80);
            Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(sheetPng.Bytes, out OfficeRasterImage? sheetRendered));
            Assert.NotNull(sheetRendered);
            Assert.True(sheetRendered!.Width >= 128);
            Assert.True(sheetRendered.Height >= 40);
        }

        [Fact]
        public void ExcelRange_ImageExportEnforcesAggregateSourceImageBudget() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Images");
            byte[] image = CreateSolidPng(12, 10, OfficeColor.FromRgb(37, 99, 235));
            sheet.AddImage(1, 1, image, "image/png", widthPixels: 12, heightPixels: 10, name: "First");
            sheet.AddImage(2, 1, image, "image/png", widthPixels: 12, heightPixels: 10, name: "Second");
            var options = new ExcelImageExportOptions {
                MaximumTotalEncodedBytes = 1,
                MaximumTotalSourceImageBytes = image.LongLength
            };

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:B3").CreateVisualSnapshot(options);

            Assert.Single(snapshot.Images);
            Assert.Contains(snapshot.Diagnostics, diagnostic =>
                diagnostic.Code == ExcelImageExportDiagnosticCodes.ImageBytesMissing &&
                diagnostic.Message.Contains("aggregate source-image budget", StringComparison.Ordinal));
        }

        [Fact]
        public void ExcelImageExportLimitsRejectOversizedSeekableSourceBeforeReading() {
            using var source = new LengthOnlyImageStream(ExcelImageExportOptions.DefaultMaximumTotalSourceImageBytes + 1L);

            bool success = ExcelImageExportLimits.TryReadSourceImageBytes(source, out byte[] bytes);

            Assert.False(success);
            Assert.Empty(bytes);
            Assert.Equal(0, source.ReadCount);
        }

        [Fact]
        public void ExcelImageExportLimitsUseTheCallerAggregateBudgetWithoutAHiddenPerImageCap() {
            const long priorPerImageLimit = 128L * 1024L * 1024L;
            using var source = new LengthOnlyImageStream(priorPerImageLimit + 1L);
            using var defaultBudgetSource = new LengthOnlyImageStream(priorPerImageLimit + 1L);

            bool success = ExcelImageExportLimits.TryReadSourceImageBytes(source, priorPerImageLimit, out byte[] bytes);

            Assert.False(success);
            Assert.Empty(bytes);
            Assert.Equal(0, source.ReadCount);

            success = ExcelImageExportLimits.TryReadSourceImageBytes(source, priorPerImageLimit + 1L, out bytes);

            Assert.False(success);
            Assert.Empty(bytes);
            Assert.Equal(1, source.ReadCount);

            success = ExcelImageExportLimits.TryReadSourceImageBytes(defaultBudgetSource, out bytes);

            Assert.False(success);
            Assert.Empty(bytes);
            Assert.Equal(1, defaultBudgetSource.ReadCount);
        }

        [Fact]
        public void ExcelRange_ImageExportIncludesAndClipsImagesOverlappingSelectedRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ImageClip");
            sheet.SetColumnWidth(1, 8);
            sheet.SetColumnWidth(2, 8);
            sheet.SetRowHeight(1, 28);
            sheet.CellValue(1, 2, "Visible");
            byte[] banner = CreateSolidPng(96, 24, OfficeColor.FromRgb(220, 38, 38));
            sheet.AddImage(1, 1, banner, "image/png", widthPixels: 96, heightPixels: 24, name: "WideBanner");

            ExcelRange range = sheet.Range("B1:B1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal(96D, image.SourceWidth);
            Assert.Equal(24D, image.SourceHeight);
            Assert.True(image.X < 0D, "The overlapping image should keep its true negative X position relative to the exported range.");
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("clip-path=\"url(#xl-image-clip-", svg, StringComparison.Ordinal);
            Assert.Contains("x=\"-", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            AssertPixelNear(rendered!, 8, 8, OfficeColor.FromRgb(220, 38, 38), tolerance: 3);
        }

        [Fact]
        public void ExcelWorksheet_ImageExportNormalizesExplicitRangeReferences() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Normalize");
            sheet.CellValue(1, 1, "Normalized");

            OfficeImageExportResult png = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions {
                Range = "'Normalize'!$A$1",
                ShowGridlines = false
            });

            Assert.Equal("Normalize!A1:A1", png.Source);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
        }

        [Fact]
        public void ExcelRange_ImageExportRendersMergedCellsStartingOutsideSelectedRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("MergedClip");
            sheet.SetColumnWidth(1, 8);
            sheet.SetColumnWidth(2, 8);
            sheet.SetColumnWidth(3, 8);
            sheet.CellValue(1, 1, "Merged title");
            sheet.Range("A1:C1").Merge();

            ExcelRangeVisualSnapshot snapshot = sheet.Range("B1:C1").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });

            ExcelVisualCell mergedOrigin = Assert.Single(snapshot.Cells, cell => cell.Row == 1 && cell.Column == 1 && !cell.CoveredByMerge);
            Assert.Equal("Merged title", mergedOrigin.Text);
            Assert.True(mergedOrigin.X < 0D, "The merged origin should keep its true negative X position relative to the selected range.");
            Assert.True(mergedOrigin.Width > snapshot.Width, "The merged origin should retain the full merged-cell width for clipping.");
        }

        [Fact]
        public void ExcelRange_ImageExportBoundsMergedRangesBeforeExpansion() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("MergedLimit");
            sheet.CellValue(1, 1, "Bounded");
            sheet.Range("A1:XFD1048576").Merge();

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:A1").CreateVisualSnapshot();

            ExcelVisualCell cell = Assert.Single(snapshot.Cells);
            Assert.Equal("Bounded", cell.Text);
            Assert.True(cell.Width <= snapshot.Width);
            Assert.True(cell.Height <= snapshot.Height);
        }

        [Fact]
        public void ExcelRange_ImageExportRetainsVisibleMergesAfterLargeOffRangeOrigin() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("MergeBudget");
            sheet.CellValue(1, 1, "Outside origin");
            sheet.CellValue(1, 3, "Visible origin");
            sheet.Range("A1:B50").Merge();
            sheet.Range("C1:C2").Merge();

            ExcelRangeVisualSnapshot snapshot = sheet.Range("B1:C2").CreateVisualSnapshot(
                new ExcelImageExportOptions { MaximumRenderedCells = 100 });

            Assert.Contains(snapshot.Cells, cell => cell.Row == 1 && cell.Column == 1 && !cell.CoveredByMerge);
            ExcelVisualCell covered = snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 3);
            Assert.True(covered.CoveredByMerge);
            Assert.True(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 3).Height > covered.Height);
        }

        [Fact]
        public void ExcelRange_ImageExportUsesTwoCellImageAnchorDimensions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            byte[] banner = CreateSolidPng(32, 32, OfficeColor.FromRgb(37, 99, 235));
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("TwoCell");
                sheet.SetColumnWidth(1, 10);
                sheet.SetColumnWidth(2, 10);
                sheet.SetColumnWidth(3, 10);
                sheet.SetColumnWidth(4, 10);
                sheet.SetRowHeight(1, 24);
                sheet.SetRowHeight(2, 24);
                sheet.SetRowHeight(3, 24);
                sheet.CellValue(1, 1, "Two-cell");
                document.Save();
            }

            AddTwoCellAnchoredImage(filePath, banner);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loaded.Sheets.Single();
            ExcelRange range = loadedSheet.Range("A1:D4");
            ExcelImageExportOptions options = new ExcelImageExportOptions { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal("TwoCell!TwoCellBanner", image.Source);
            Assert.True(image.Width > 150D, $"Expected two-cell anchor width, got {image.Width}.");
            Assert.True(image.Height > 40D, $"Expected two-cell anchor height, got {image.Height}.");
            Assert.Contains("data:image/png;base64,", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            AssertPixelNear(rendered!, Math.Min(rendered!.Width - 1, 120), Math.Min(rendered.Height - 1, 40), OfficeColor.FromRgb(37, 99, 235), tolerance: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportSaturatesOverflowingTwoCellAnchorOffsets() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            byte[] banner = CreateSolidPng(8, 8, OfficeColor.FromRgb(37, 99, 235));
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("AnchorOverflow");
                sheet.CellValue(1, 1, "Anchor");
                document.Save();
            }

            AddTwoCellAnchoredImage(
                filePath,
                banner,
                fromColumnOffset: long.MinValue.ToString(CultureInfo.InvariantCulture),
                toColumnOffset: long.MaxValue.ToString(CultureInfo.InvariantCulture),
                toColumnId: "0");

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelImage image = Assert.Single(loaded.Sheets.Single().Images);

            Assert.Equal(16384, image.WidthPixels);
        }

        [Fact]
        public void ExcelRange_ImageExportRespectsAbsoluteAnchorImageCoordinates() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            byte[] banner = CreateSolidPng(80, 20, OfficeColor.FromRgb(22, 163, 74));
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("AbsoluteImage");
                sheet.SetColumnWidth(1, 8);
                sheet.SetColumnWidth(2, 8);
                sheet.SetRowHeight(1, 24);
                document.Save();
            }

            AddAbsoluteAnchoredImage(filePath, banner, xPixels: 40, yPixels: 0, widthPixels: 80, heightPixels: 20);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loaded.Sheets.Single();
            ExcelRangeVisualSnapshot snapshot = loadedSheet.Range("B1:C2").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal("AbsoluteImage!AbsoluteBanner", image.Source);
            Assert.True(image.X < 0D, "Absolute-anchor images should keep their worksheet-canvas X position relative to the selected range.");
            Assert.Equal(80D, image.Width);
            Assert.Equal(20D, image.Height);
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsPictureCropRectangle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            byte[] croppedSource = CreateHorizontalCropPng();
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("Crop");
                sheet.SetColumnWidth(1, 14);
                sheet.SetColumnWidth(2, 14);
                sheet.SetRowHeight(1, 30);
                sheet.SetRowHeight(2, 30);
                document.Save();
            }

            AddCroppedImage(filePath, croppedSource);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loaded.Sheets.Single();
            ExcelRange range = loadedSheet.Range("A1:B2");
            ExcelImageExportOptions options = new ExcelImageExportOptions { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal(0.25D, image.CropLeftRatio, precision: 3);
            Assert.Equal(0.25D, image.CropRightRatio, precision: 3);
            Assert.True(image.SourceWidth > 0D);
            Assert.True(image.SourceHeight > 0D);
            Assert.True(image.HasCrop);
            Assert.Contains("clip-path=\"url(#xl-image-clip-", svg, StringComparison.Ordinal);
            Assert.Contains("x=\"-", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            AssertPixelNear(rendered!, 8, 8, OfficeColor.FromRgb(37, 99, 235), tolerance: 8);
            AssertPixelNear(rendered!, Math.Min(rendered.Width - 1, 90), 8, OfficeColor.FromRgb(37, 99, 235), tolerance: 8);
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsPictureRotation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            byte[] rotatedSource = CreateRotationProbePng();
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("RotateImage");
                for (int column = 1; column <= 4; column++) {
                    sheet.SetColumnWidth(column, 14);
                }

                for (int row = 1; row <= 5; row++) {
                    sheet.SetRowHeight(row, 30);
                }

                document.Save();
            }

            AddRotatedImage(filePath, rotatedSource);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loaded.Sheets.Single();
            ExcelRange range = loadedSheet.Range("A1:D5");
            ExcelImageExportOptions options = new ExcelImageExportOptions {
                ShowGridlines = false,
                BackgroundColor = OfficeColor.White
            };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal("RotateImage!RotatedBanner", image.Source);
            Assert.Equal(30D, image.RotationDegrees, precision: 3);
            Assert.True(image.HasRotation);
            Assert.Contains("transform=\"rotate(30", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsVisiblePixel(
                    rendered!,
                    image.X - 8D,
                    image.Y - 32D,
                    image.X + image.Width,
                    image.Y - 1D),
                "Expected rotated picture pixels above the unrotated image rectangle.");
        }

        [Fact]
        public void ExcelRange_ImageExportCombinesPictureCropRotationAndFlip() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            byte[] transformedSource = CreateTransformProbePng();
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("TransformImage");
                for (int column = 1; column <= 5; column++) {
                    sheet.SetColumnWidth(column, 14);
                }

                for (int row = 1; row <= 6; row++) {
                    sheet.SetRowHeight(row, 30);
                }

                document.Save();
            }

            AddTransformedImage(filePath, transformedSource);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            ExcelSheet loadedSheet = loaded.Sheets.Single();
            ExcelRange range = loadedSheet.Range("A1:E6");
            ExcelImageExportOptions options = new ExcelImageExportOptions {
                ShowGridlines = false,
                BackgroundColor = OfficeColor.White
            };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal("TransformImage!TransformedBanner", image.Source);
            Assert.True(image.HasCrop);
            Assert.True(image.HasRotation);
            Assert.True(image.FlipHorizontal);
            Assert.Equal(0.25D, image.CropLeftRatio, precision: 3);
            Assert.Equal(30D, image.RotationDegrees, precision: 3);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == "ExcelImageFlipUnsupported");
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == "ExcelImageCropRotationCombinationUnsupported");
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("rotate(30", svg, StringComparison.Ordinal);
            Assert.Contains("scale(-1 1)", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsVisiblePixel(
                    rendered!,
                    image.X - 8D,
                    image.Y - 32D,
                    image.X + image.Width,
                    image.Y - 1D),
                "Expected cropped, flipped, rotated picture pixels above the unrotated image rectangle.");
        }

        [Fact]
        public void ExcelRange_ImageExportEmbedsJpegInSvgAndUsesVisibleRasterFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Jpeg");
            sheet.CellValue(1, 1, "Photo");
            sheet.AddImage(1, 2, CreateMinimalJpegHeader(), "image/jpeg", widthPixels: 16, heightPixels: 12, name: "PhotoJpeg");

            ExcelRange range = sheet.Range("A1:C3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal(OfficeImageFormat.Jpeg, image.DetectedFormat);
            Assert.Contains("data:image/jpeg;base64,", svgText, StringComparison.Ordinal);
            Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Jpeg!PhotoJpeg", diagnostic.Source);
        }

        [Fact]
        public void ExcelRange_ImageExportRendersBmpThroughSharedRasterDecoder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Bitmap");
            sheet.CellValue(1, 1, "Bitmap");
            OfficeColor color = OfficeColor.FromRgb(18, 52, 86);
            sheet.AddImage(1, 2, CreateBmp24(2, 2, new[] { color, color, color, color }), "image/bmp", widthPixels: 24, heightPixels: 18, name: "BitmapOverlay");

            ExcelRange range = sheet.Range("A1:C3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg);

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal(OfficeImageFormat.Bmp, image.DetectedFormat);
            Assert.Equal("image/bmp", image.ContentType);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.True(
                ContainsPixelNear(rendered!, image.X, image.Y, image.X + image.Width, image.Y + image.Height, color, tolerance: 8),
                "Expected Excel PNG export to contain decoded BMP pixels from the shared raster decoder.");
        }

        [Fact]
        public void ExcelRange_ImageExportRendersTopDownBmpThroughSharedRasterDecoder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("TopDownBitmap");
            sheet.CellValue(1, 1, "Top-down bitmap");
            OfficeColor color = OfficeColor.FromRgb(24, 96, 144);
            sheet.AddImage(1, 2, CreateBmp24(2, 2, new[] { color, color, color, color }, topDown: true), "image/bmp", widthPixels: 24, heightPixels: 18, name: "TopDownBitmapOverlay");

            ExcelRange range = sheet.Range("A1:C3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg);

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal(OfficeImageFormat.Bmp, image.DetectedFormat);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.True(
                ContainsPixelNear(rendered!, image.X, image.Y, image.X + image.Width, image.Y + image.Height, color, tolerance: 8),
                "Expected Excel PNG export to contain decoded top-down BMP pixels from the shared raster decoder.");
        }

        [Fact]
        public void ExcelRange_ImageExportRendersBmp32AlphaThroughSharedRasterDecoder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("BitmapAlpha");
            sheet.CellValue(1, 1, "Bitmap alpha");
            byte[] bmp = CreateBmp32(2, 2, new[] {
                OfficeColor.FromRgba(255, 0, 0, 128), OfficeColor.FromRgba(255, 0, 0, 128),
                OfficeColor.FromRgba(255, 0, 0, 128), OfficeColor.FromRgba(255, 0, 0, 128)
            });
            sheet.AddImage(1, 2, bmp, "image/bmp", widthPixels: 24, heightPixels: 18, name: "BitmapAlphaOverlay");

            ExcelRange range = sheet.Range("A1:C3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg);

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal(OfficeImageFormat.Bmp, image.DetectedFormat);
            Assert.Equal("image/bmp", image.ContentType);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/png;base64,", svgText, StringComparison.Ordinal);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.True(
                ContainsPixelNear(rendered!, image.X, image.Y, image.X + image.Width, image.Y + image.Height, OfficeColor.FromRgb(255, 127, 127), tolerance: 8),
                "Expected Excel PNG export to contain alpha-blended BMP32 pixels from the shared raster decoder.");
        }

        [Fact]
        public void ExcelRange_ImageExportRendersGifThroughSharedRasterDecoder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Gif");
            sheet.CellValue(1, 1, "GIF");
            sheet.AddImage(1, 2, CreateSinglePixelGif(), "image/gif", widthPixels: 24, heightPixels: 18, name: "GifOverlay");

            ExcelRange range = sheet.Range("A1:C3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg);

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal(OfficeImageFormat.Gif, image.DetectedFormat);
            Assert.Equal("image/gif", image.ContentType);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            string svgText = System.Text.Encoding.UTF8.GetString(svg.Bytes);
            Assert.Contains("data:image/gif;base64,", svgText, StringComparison.Ordinal);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.True(
                ContainsPixelNear(rendered!, image.X, image.Y, image.X + image.Width, image.Y + image.Height, OfficeColor.White, tolerance: 2),
                "Expected Excel PNG export to contain decoded GIF pixels from the shared raster decoder.");
        }

        [Fact]
        public void ExcelWorksheet_ImageExportExpandsPngRangeForRasterDecodableBmpImages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("WorksheetBitmap");
            sheet.CellValue(1, 1, "Used cell");
            OfficeColor color = OfficeColor.FromRgb(18, 52, 86);
            sheet.AddImage(10, 5, CreateBmp24(2, 2, new[] { color, color, color, color }), "image/bmp", widthPixels: 24, heightPixels: 18, name: "BitmapOutsideUsedCell");

            OfficeImageExportResult png = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions { ShowGridlines = false });

            Assert.Equal("WorksheetBitmap!A1:E10", png.Source);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(ContainsPixelNear(rendered!, 0D, 0D, rendered.Width, rendered.Height, color, tolerance: 8));
        }

        [Fact]
        public void ExcelWorksheet_ImageExportExpandsPngRangeForIdentifiedUndecodablePngImages() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("WorksheetPng");
            sheet.CellValue(1, 1, "Used cell");
            sheet.AddImage(10, 5, CreateTruncatedPngHeader(4, 3), "image/png", widthPixels: 24, heightPixels: 18, name: "TruncatedPngOutsideUsedCell");

            OfficeImageExportResult png = sheet.ExportImage(OfficeImageExportFormat.Png, new ExcelWorksheetImageExportOptions { ShowGridlines = false });

            Assert.Equal("WorksheetPng!A1:E10", png.Source);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.Equal("WorksheetPng!TruncatedPngOutsideUsedCell", diagnostic.Source);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnknownImageFormatWithStableCodeAndSource() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("UnknownImage");
            sheet.CellValue(1, 1, "Image");
            sheet.AddImage(1, 2, new byte[] { 1, 2, 3, 4, 5, 6 }, "image/bmp", widthPixels: 16, heightPixels: 12, name: "MysteryBitmap");

            ExcelRange range = sheet.Range("A1:C3");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult svg = range.ExportImage(OfficeImageExportFormat.Svg);

            ExcelVisualImage image = Assert.Single(snapshot.Images);
            Assert.Equal(OfficeImageFormat.Unknown, image.DetectedFormat);
            Assert.Equal("image/bmp", image.ContentType);

            OfficeImageExportDiagnostic snapshotDiagnostic = Assert.Single(snapshot.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ImageFormatUnknown);
            Assert.Equal("UnknownImage!MysteryBitmap", snapshotDiagnostic.Source);
            Assert.Contains("image/bmp", snapshotDiagnostic.Message, StringComparison.Ordinal);

            OfficeImageExportDiagnostic pngDiagnostic = Assert.Single(png.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.Equal("UnknownImage!MysteryBitmap", pngDiagnostic.Source);
            Assert.Contains("image/bmp", pngDiagnostic.Message, StringComparison.Ordinal);

            OfficeImageExportDiagnostic svgDiagnostic = Assert.Single(svg.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.Equal("UnknownImage!MysteryBitmap", svgDiagnostic.Source);
            Assert.Contains("image/bmp", svgDiagnostic.Message, StringComparison.Ordinal);
        }

        [Theory]
        [InlineData(OfficeImageExportFormat.Png)]
        [InlineData(OfficeImageExportFormat.Svg)]
        public void ExcelWorksheet_ImageExportUsesCallerCodecAndExpandsImageRange(OfficeImageExportFormat format) {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("CallerCodec");
            sheet.CellValue(1, 1, "Used cell");
            sheet.AddImage(
                10,
                5,
                new byte[] { 1, 2, 3, 4 },
                "image/tiff",
                widthPixels: 24,
                heightPixels: 18,
                name: "CustomOutsideUsedCell");
            var codec = new SolidImageCodec(OfficeColor.FromRgb(18, 86, 140));
            var options = new ExcelWorksheetImageExportOptions {
                ShowGridlines = false,
                ImageCodec = codec
            };

            OfficeImageExportResult result = sheet.ExportImage(format, options);

            Assert.Equal("CallerCodec!A1:E10", result.Source);
            Assert.Equal(1, codec.DecodeCalls);
            Assert.Contains(
                result.Diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec &&
                              diagnostic.Source == "CallerCodec!CustomOutsideUsedCell");
            Assert.DoesNotContain(
                result.Diagnostics,
                diagnostic => diagnostic.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            if (format == OfficeImageExportFormat.Svg) {
                Assert.Contains("data:image/png;base64,", System.Text.Encoding.UTF8.GetString(result.Bytes), StringComparison.Ordinal);
            }
        }

        [Fact]
        public void ExcelRange_ExportsChartsThroughSharedDrawingRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Charts");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Revenue");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 220, heightPixels: 120, type: ExcelChartType.ColumnClustered, title: "Revenue Trend");

            ExcelRange range = sheet.Range("A1:G8");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png);
            OfficeImageExportResult sheetPng = sheet.ExportImage(OfficeImageExportFormat.Png);
            string svg = range.ToSvg();

            Assert.Single(snapshot.Charts);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code.StartsWith("ExcelChart", StringComparison.Ordinal));
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(rendered!.Width >= 448);
            Assert.True(rendered.Height >= 160);
            Assert.Contains("Revenue Trend", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(sheetPng.Bytes, out OfficeRasterImage? sheetRendered));
            Assert.NotNull(sheetRendered);
            Assert.True(sheetRendered!.Width >= 448);
            Assert.True(sheetRendered.Height >= 120);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartStyleAndDataLabelsIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartStyle");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Revenue");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 260, heightPixels: 160, type: ExcelChartType.ColumnClustered, title: "Revenue Trend");
            chart.SetChartAreaStyle(fillColor: "F2F2F2", lineColor: "404040", lineWidthPoints: 1);
            chart.SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "00B0F0", lineWidthPoints: 1D);
            chart.SetLegendTextStyle(bold: true, fontName: "Aptos Legend");
            chart.SetCategoryAxisTitle("Quarter")
                .SetValueAxisTitle("Revenue")
                .SetCategoryAxisLabelTextStyle(italic: true, fontName: "Aptos Axis Label")
                .SetValueAxisLabelTextStyle(italic: true, fontName: "Aptos Axis Label")
                .SetCategoryAxisTitleTextStyle(bold: true, fontName: "Aptos Axis Title")
                .SetValueAxisTitleTextStyle(bold: true, fontName: "Aptos Axis Title");
            chart.SetDataLabels(
                showValue: true,
                showCategoryName: true,
                showSeriesName: false,
                showLegendKey: false,
                showPercent: false,
                position: C.DataLabelPositionValues.OutsideEnd,
                numberFormat: "0");
            chart.SetDataLabelTextStyle(italic: true, fontName: "Aptos Data");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(OfficeColor.FromRgb(242, 242, 242), visualChart.Snapshot.Style!.BackgroundColor);
            Assert.Equal(OfficeColor.FromRgb(64, 64, 64), visualChart.Snapshot.Style.BorderColor);
            Assert.Equal(OfficeColor.White, visualChart.Snapshot.Style.PlotAreaBackgroundColor);
            Assert.Equal(OfficeColor.FromRgb(0, 176, 240), visualChart.Snapshot.Style.PlotAreaBorderColor);
            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.True(visualChart.Snapshot.Layout!.ShowDataLabels);
            Assert.True(visualChart.Snapshot.Layout.ShowDataLabelValues);
            Assert.True(visualChart.Snapshot.Layout.ShowDataLabelCategoryNames);
            Assert.Equal(OfficeChartDataLabelPosition.OutsideEnd, visualChart.Snapshot.Layout.DataLabelPosition);
            Assert.Equal("0", visualChart.Snapshot.Layout.DataLabelNumberFormat);
            Assert.Equal("Aptos Legend", visualChart.Snapshot.Layout.LegendFontFamily);
            Assert.Equal("Aptos Data", visualChart.Snapshot.Layout.DataLabelFontFamily);
            Assert.Equal("Aptos Axis Label", visualChart.Snapshot.Layout.AxisTextFontFamily);
            Assert.Equal("Aptos Axis Title", visualChart.Snapshot.Layout.AxisTitleFontFamily);
            Assert.Equal(OfficeFontStyle.Bold, visualChart.Snapshot.Layout.LegendFontStyle);
            Assert.Equal(OfficeFontStyle.Italic, visualChart.Snapshot.Layout.DataLabelFontStyle);
            Assert.Equal(OfficeFontStyle.Italic, visualChart.Snapshot.Layout.AxisTextFontStyle);
            Assert.Equal(OfficeFontStyle.Bold, visualChart.Snapshot.Layout.AxisTitleFontStyle);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#F2F2F2", svg, StringComparison.Ordinal);
            Assert.Contains("#00B0F0", svg, StringComparison.Ordinal);
            Assert.Contains("font-family=\"Aptos Legend\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-family=\"Aptos Data\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-family=\"Aptos Axis Label\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-family=\"Aptos Axis Title\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("Jan; 120", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            AssertPixelNear(
                rendered!,
                (int)((visualChart.X + 5D) * options.Scale),
                (int)((visualChart.Y + 5D) * options.Scale),
                OfficeColor.FromRgb(242, 242, 242),
                tolerance: 4);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(0, 176, 240),
                    tolerance: 8),
                "Expected the exported chart to include the authored plot area border color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartAndPlotAreaBorderWidthAndDashIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartAreaLines");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Revenue");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 260, heightPixels: 160, type: ExcelChartType.ColumnClustered, title: "Area Lines");
            chart.SetChartAreaStyle(fillColor: "F2F2F2", lineColor: "404040", lineWidthPoints: 3D);
            chart.SetPlotAreaStyle(fillColor: "FFFFFF", lineColor: "00B0F0", lineWidthPoints: 2D);
            SetFirstChartAreaDash(document, A.PresetLineDashValues.DashDot);
            SetFirstChartPlotAreaDash(document, A.PresetLineDashValues.Dot);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(3D, visualChart.Snapshot.Style!.ChartBorderWidth);
            Assert.Equal(2D, visualChart.Snapshot.Style.PlotAreaBorderWidth);
            Assert.Equal(OfficeStrokeDashStyle.DashDot, visualChart.Snapshot.Style.ChartBorderDashStyle);
            Assert.Equal(OfficeStrokeDashStyle.Dot, visualChart.Snapshot.Style.PlotAreaBorderDashStyle);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartAreaStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("stroke=\"#404040\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke=\"#00B0F0\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-width=\"3\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-width=\"2\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray=\"12 6 3 6\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray=\"2 4\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(0, 176, 240),
                    tolerance: 8),
                "Expected the exported chart to include the authored dashed plot area border color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartTitleColorIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartTitle");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Title Color");
            chart.SetTitleTextStyle(color: "BE123C");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 3D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(OfficeColor.FromRgb(190, 18, 60), visualChart.Snapshot.Style!.TitleColor);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#BE123C", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(190, 18, 60),
                    tolerance: 40),
                "Expected the exported chart to include the authored title color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartTitleTypographyIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartTitleTypography");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Title Typography");
            chart.SetTitleTextStyle(fontSizePoints: 16D, bold: false, italic: true, color: "BE123C", fontName: "Aptos Display");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(OfficeColor.FromRgb(190, 18, 60), visualChart.Snapshot.Style!.TitleColor);
            Assert.Equal("Aptos Display", visualChart.Snapshot.Style.TitleFontFamily);
            Assert.Equal(16D, visualChart.Snapshot.Style.TitleFontSize);
            Assert.Equal(OfficeFontStyle.Italic, visualChart.Snapshot.Style.TitleFontStyle);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("font-family=\"Aptos Display\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-size=\"16\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.OrdinalIgnoreCase);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnsupportedChartTitleTextStyle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartTitleStyle");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Title Style");
            chart.SetTitleTextStyle(color: "BE123C");
            AddFirstChartTitleTextEffect(document);

            OfficeImageExportResult png = sheet.Range("A1:H9").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartTextStyleApproximation);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("ChartTitleStyle!" + chart.Name, diagnostic.Source);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartSeriesColorsIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartSeries");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(1, 3, "Forecast");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(2, 3, 80);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(3, 3, 110);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            sheet.CellValue(4, 3, 130);
            ExcelChart chart = sheet.AddChartFromRange("A1:C4", row: 1, column: 5, widthPixels: 280, heightPixels: 170, type: ExcelChartType.ColumnClustered, title: "Series Colors");
            chart.SetSeriesFillColor(0, "22C55E");
            chart.SetSeriesFillColor(1, "F97316");

            ExcelRange range = sheet.Range("A1:I9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("22C55E", visualChart.Snapshot.Data.Series[0].SeriesColorArgb);
            Assert.Equal("F97316", visualChart.Snapshot.Data.Series[1].SeriesColorArgb);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#22C55E", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#F97316", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(34, 197, 94),
                    tolerance: 8),
                "Expected the exported chart to include the authored first series color.");
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(249, 115, 22),
                    tolerance: 8),
                "Expected the exported chart to include the authored second series color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartPointColorsIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartPoints");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Point Colors");
            chart.SetSeriesFillColor(0, "22C55E");
            chart.SetSeriesPointFillColor(0, 1, "A855F7");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Data.Series[0].PointColorArgb);
            Assert.Equal("A855F7", visualChart.Snapshot.Data.Series[0].PointColorArgb![1]);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#A855F7", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(168, 85, 247),
                    tolerance: 8),
                "Expected the exported chart to include the authored point color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesSimpleChartMarkerFillIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartMarkers");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "Marker Fill");
            chart.SetSeriesLineColor(0, "2563EB");
            chart.SetSeriesMarker(0, C.MarkerStyleValues.Circle, fillColor: "F97316");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 3D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.True(visualChart.Snapshot.Data.Series[0].ShowMarkers);
            Assert.NotNull(visualChart.Snapshot.Data.Series[0].PointColorArgb);
            Assert.All(visualChart.Snapshot.Data.Series[0].PointColorArgb!, color => Assert.Equal("F97316", color));
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#F97316", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(249, 115, 22),
                    tolerance: 8),
                "Expected the exported chart to include the authored marker fill color.");
        }

        [Fact]
        public void ExcelRange_ImageExportDoesNotInventLineChartMarkersWhenSourceHasNoMarkerElement() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("NoMarkers");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "No Markers");
            chart.SetSeriesLineColor(0, "2563EB");
            GetFirstChartPart(document).ChartSpace.Descendants<C.LineChartSeries>().First().RemoveAllChildren<C.Marker>();

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:H9").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);

            Assert.False(visualChart.Snapshot.Data.Series[0].ShowMarkers);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartSeriesNoLineIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("NoSeriesLine");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "No Series Line");
            chart.SetSeriesLineColor(0, "2563EB");
            SetFirstChartSeriesNoLine(document);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:H9").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);

            Assert.False(visualChart.Snapshot.Data.Series[0].ConnectLine);
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesChartSeriesStyleByElementOrderWhenIndexesAreNonContiguous() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("SeriesOrder");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "Series Order");
            chart.SetSeriesLineColor(0, "2563EB");
            SetFirstChartSeriesIndex(document, 1U);

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:H9").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);

            Assert.Equal("2563EB", visualChart.Snapshot.Data.Series[0].SeriesColorArgb);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsFormulaConditionalFormatThresholds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("FormulaThresholds");
            for (int row = 1; row <= 3; row++) {
                sheet.CellValue(row, 1, row * 10);
                sheet.CellValue(row, 2, row * 10);
            }

            sheet.AddConditionalDataBar("A1:A3", OfficeColor.Blue);
            sheet.AddConditionalColorScale("B1:B3", OfficeColor.Red, OfficeColor.Green);
            SetFirstDataBarThresholdFormula(sheet, "A1");
            SetFirstColorScaleThresholdFormula(sheet, "B1");

            OfficeImageExportResult png = sheet.Range("A1:B3").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(
                2,
                png.Diagnostics.Count(diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalFormulaThresholdApproximation));
        }

        [Fact]
        public void ExcelRange_ImageExportResolvesAbsoluteAnchorChartCoordinates() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("AbsoluteChart");
                sheet.CellValue(1, 1, "Month");
                sheet.CellValue(1, 2, "Actual");
                sheet.CellValue(2, 1, "Jan");
                sheet.CellValue(2, 2, 120);
                sheet.CellValue(3, 1, "Feb");
                sheet.CellValue(3, 2, 180);
                sheet.AddChartFromRange("A1:B3", row: 4, column: 6, widthPixels: 260, heightPixels: 160, type: ExcelChartType.ColumnClustered, title: "Absolute Chart");
                document.Save();
            }

            MoveFirstChartToAbsoluteAnchor(filePath, xPixels: 40, yPixels: 30, widthPixels: 220, heightPixels: 120);

            using (ExcelDocument document = ExcelDocument.Load(filePath)) {
                ExcelSheet sheet = document.Sheets.Single();
                ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:J12").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
                ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);

                Assert.Equal(40D, visualChart.X);
                Assert.Equal(30D, visualChart.Y);
                Assert.Equal(220D, visualChart.Width);
                Assert.Equal(120D, visualChart.Height);
            }
        }

        [Fact]
        public void ExcelWorksheet_ImageExportRejectsDrawingExpandedRangesBeforeMaterialization() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorksheet("BoundedDrawingRange");
                sheet.CellValue(1, 1, "Category");
                sheet.CellValue(1, 2, "Value");
                sheet.CellValue(2, 1, "Item");
                sheet.CellValue(2, 2, 1);
                sheet.AddChartFromRange("A1:B2", row: 1, column: 3, widthPixels: 240, heightPixels: 140, type: ExcelChartType.ColumnClustered);
                document.Save();
            }

            MoveFirstChartToAbsoluteAnchor(filePath, xPixels: 1_000_000, yPixels: 1_000_000, widthPixels: 1_000_000, heightPixels: 1_000_000);

            using ExcelDocument loaded = ExcelDocument.Load(filePath);
            InvalidOperationException error = Assert.Throws<InvalidOperationException>(() =>
                loaded.Sheets.Single().ExportImage(
                    OfficeImageExportFormat.Png,
                    new ExcelWorksheetImageExportOptions { ShowGridlines = false }));

            Assert.Contains("100000 rendered cells", error.Message, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesStandardRadarChartsWithoutFill() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("RadarStandard");
            sheet.CellValue(1, 1, "Metric");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Quality");
            sheet.CellValue(2, 2, 7);
            sheet.CellValue(3, 1, "Speed");
            sheet.CellValue(3, 2, 5);
            sheet.CellValue(4, 1, "Cost");
            sheet.CellValue(4, 2, 8);
            sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Radar, title: "Radar");

            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:H9").CreateVisualSnapshot(new ExcelImageExportOptions { ShowGridlines = false });
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.False(visualChart.Snapshot.Layout!.FillRadarSeries);
        }

        [Fact]
        public void ExcelRange_ImageExportKeepsNoFillChartsTransparentInPng() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("TransparentChart");
            for (int row = 1; row <= 10; row++) {
                for (int column = 1; column <= 8; column++) {
                    sheet.CellValue(row, column, row + column);
                }
            }

            sheet.Range("A1:H10").SetFillColor("22C55E");
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "Transparent");
            chart.SetPlotAreaStyle(fillColor: "FFFFFF");
            SetFirstChartAreaNoFill(document);

            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = sheet.Range("A1:H10").CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = sheet.Range("A1:H10").ExportImage(OfficeImageExportFormat.Png, options);

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            AssertPixelNear(
                rendered!,
                (int)((visualChart.X + 4D) * options.Scale),
                (int)((visualChart.Y + 4D) * options.Scale),
                OfficeColor.FromRgb(34, 197, 94),
                tolerance: 8);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartMarkerSizeIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("MarkerSize");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "Marker Size");
            chart.SetSeriesLineColor(0, "2563EB");
            chart.SetSeriesMarker(0, C.MarkerStyleValues.Circle, size: 12, fillColor: "F97316");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal(12, visualChart.Snapshot.Data.Series[0].MarkerSize);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("rx=\"6\" ry=\"6\"", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(249, 115, 22),
                    tolerance: 8),
                "Expected the exported chart to include the authored marker fill color at the authored marker size.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartMarkerShapeIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("MarkerShape");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "Marker Shape");
            chart.SetSeriesLineColor(0, "2563EB");
            chart.SetSeriesMarker(0, C.MarkerStyleValues.Diamond, size: 12, fillColor: "F97316");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal(OfficeChartMarkerShape.Diamond, visualChart.Snapshot.Data.Series[0].MarkerShape);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("<polygon", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#F97316", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(249, 115, 22),
                    tolerance: 8),
                "Expected the exported chart to include the authored diamond marker fill color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartXMarkerShapeIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("MarkerX");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "Marker X");
            chart.SetSeriesLineColor(0, "2563EB");
            chart.SetSeriesMarker(0, C.MarkerStyleValues.X, size: 14, lineColor: "7C3AED", lineWidthPoints: 2D);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal(OfficeChartMarkerShape.X, visualChart.Snapshot.Data.Series[0].MarkerShape);
            Assert.Equal("7C3AED", visualChart.Snapshot.Data.Series[0].MarkerOutlineColorArgb);
            Assert.Equal(2D, visualChart.Snapshot.Data.Series[0].MarkerOutlineWidth);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("<line", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke=\"#7C3AED\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(124, 58, 237),
                    tolerance: 8),
                "Expected the exported chart to include the authored X marker stroke color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesDashDotAndStarChartMarkerShapesIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("MarkerSymbols");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Dash");
            sheet.CellValue(1, 3, "Dot");
            sheet.CellValue(1, 4, "Star");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(2, 3, 140);
            sheet.CellValue(2, 4, 110);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(3, 3, 160);
            sheet.CellValue(3, 4, 170);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            sheet.CellValue(4, 3, 130);
            sheet.CellValue(4, 4, 150);
            ExcelChart chart = sheet.AddChartFromRange("A1:D4", row: 1, column: 6, widthPixels: 310, heightPixels: 185, type: ExcelChartType.Line, title: "Marker Symbols");
            chart.SetSeriesLineColor(0, "CBD5E1");
            chart.SetSeriesLineColor(1, "CBD5E1");
            chart.SetSeriesLineColor(2, "CBD5E1");
            chart.SetSeriesMarker(0, C.MarkerStyleValues.Dash, size: 14, lineColor: "DC2626", lineWidthPoints: 2D);
            chart.SetSeriesMarker(1, C.MarkerStyleValues.Dot, size: 14, fillColor: "059669");
            chart.SetSeriesMarker(2, C.MarkerStyleValues.Star, size: 14, fillColor: "F59E0B");

            ExcelRange range = sheet.Range("A1:K10");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal(OfficeChartMarkerShape.Dash, visualChart.Snapshot.Data.Series[0].MarkerShape);
            Assert.Equal(OfficeChartMarkerShape.Dot, visualChart.Snapshot.Data.Series[1].MarkerShape);
            Assert.Equal(OfficeChartMarkerShape.Star, visualChart.Snapshot.Data.Series[2].MarkerShape);
            Assert.Equal("DC2626", visualChart.Snapshot.Data.Series[0].MarkerOutlineColorArgb);
            Assert.Equal("059669", visualChart.Snapshot.Data.Series[1].PointColorArgb![0]);
            Assert.Equal("F59E0B", visualChart.Snapshot.Data.Series[2].PointColorArgb![0]);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("<line", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<ellipse", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("<polygon", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke=\"#DC2626\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#059669", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#F59E0B", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(220, 38, 38),
                    tolerance: 8),
                "Expected the exported chart to include the authored dash marker stroke color.");
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(5, 150, 105),
                    tolerance: 8),
                "Expected the exported chart to include the authored dot marker fill color.");
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(245, 158, 11),
                    tolerance: 8),
                "Expected the exported chart to include the authored star marker fill color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartMarkerOutlineIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("MarkerOutline");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "Marker Outline");
            chart.SetSeriesLineColor(0, "2563EB");
            chart.SetSeriesMarker(0, C.MarkerStyleValues.Circle, size: 16, fillColor: "F97316", lineColor: "7C2D12", lineWidthPoints: 3D);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("7C2D12", visualChart.Snapshot.Data.Series[0].MarkerOutlineColorArgb);
            Assert.Equal(3D, visualChart.Snapshot.Data.Series[0].MarkerOutlineWidth);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("stroke=\"#7C2D12\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-width=\"3\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(124, 45, 18),
                    tolerance: 8),
                "Expected the exported chart to include the authored marker outline color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartSeriesLineWidthIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("SeriesLineWidth");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "Series Width");
            chart.SetSeriesLineColor(0, "2563EB", widthPoints: 4D);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("2563EB", visualChart.Snapshot.Data.Series[0].SeriesColorArgb);
            Assert.Equal(4D, visualChart.Snapshot.Data.Series[0].SeriesLineWidth);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("stroke=\"#2563EB\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-width=\"4\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(37, 99, 235),
                    tolerance: 8),
                "Expected the exported chart to include the authored thick series line color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartSeriesLineDashIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("SeriesLineDash");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.Line, title: "Series Dash");
            chart.SetSeriesLineColor(0, "2563EB", widthPoints: 4D);
            SetFirstChartSeriesDash(document, A.PresetLineDashValues.Dash);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.Equal("2563EB", visualChart.Snapshot.Data.Series[0].SeriesColorArgb);
            Assert.Equal(4D, visualChart.Snapshot.Data.Series[0].SeriesLineWidth);
            Assert.Equal(OfficeStrokeDashStyle.Dash, visualChart.Snapshot.Data.Series[0].SeriesLineDashStyle);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartSeriesStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("stroke=\"#2563EB\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-width=\"4\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray=\"16 8\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(37, 99, 235),
                    tolerance: 8),
                "Expected the exported chart to include the authored dashed series line color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesCategoryAndValueChartGridlineColorsIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ChartGridlines");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Gridlines");
            chart.SetCategoryAxisGridlines(showMajor: true, showMinor: false, lineColor: "DC2626");
            chart.SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: "A855F7");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.True(visualChart.Snapshot.Style!.ShowGridLines);
            Assert.Equal(true, visualChart.Snapshot.Style.ShowCategoryGridLines);
            Assert.Equal(true, visualChart.Snapshot.Style.ShowValueGridLines);
            Assert.Equal(OfficeColor.FromRgb(220, 38, 38), visualChart.Snapshot.Style.CategoryGridLineColor);
            Assert.Equal(OfficeColor.FromRgb(168, 85, 247), visualChart.Snapshot.Style.ValueGridLineColor);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartGridlineStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#DC2626", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#A855F7", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            bool hasCategoryGridlinePixel = ContainsPixelNear(
                rendered!,
                visualChart.X * options.Scale,
                visualChart.Y * options.Scale,
                (visualChart.X + visualChart.Width) * options.Scale,
                (visualChart.Y + visualChart.Height) * options.Scale,
                OfficeColor.FromRgb(220, 38, 38),
                tolerance: 8)
                || ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(234, 147, 147),
                    tolerance: 36);
            bool hasValueGridlinePixel = ContainsPixelNear(
                rendered!,
                visualChart.X * options.Scale,
                visualChart.Y * options.Scale,
                (visualChart.X + visualChart.Width) * options.Scale,
                (visualChart.Y + visualChart.Height) * options.Scale,
                OfficeColor.FromRgb(168, 85, 247),
                tolerance: 8)
                || ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(211, 170, 251),
                    tolerance: 36);
            Assert.True(
                hasCategoryGridlinePixel,
                "Expected the exported chart to include the authored category gridline color.");
            Assert.True(
                hasValueGridlinePixel,
                "Expected the exported chart to include the authored value gridline color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartMinorGridlinesIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("MinorGridlines");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Minor Gridlines");
            chart.SetValueAxisGridlines(showMajor: false, showMinor: true, lineColor: "14B8A6", lineWidthPoints: 1.5D);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(false, visualChart.Snapshot.Style!.ShowValueGridLines);
            Assert.Equal(true, visualChart.Snapshot.Style.ShowValueMinorGridLines);
            Assert.Equal(OfficeColor.FromRgb(20, 184, 166), visualChart.Snapshot.Style.ValueMinorGridLineColor);
            Assert.Equal(1.5D, visualChart.Snapshot.Style.ValueMinorGridLineWidth);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartGridlineStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("stroke=\"#14B8A6\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-width=\"1.5\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(20, 184, 166),
                    tolerance: 12),
                "Expected the exported chart to include the authored minor gridline color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartGridlineWidthIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("GridlineWidth");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Gridline Width");
            chart.SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: "A855F7", lineWidthPoints: 2D);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(true, visualChart.Snapshot.Style!.ShowValueGridLines);
            Assert.Equal(OfficeColor.FromRgb(168, 85, 247), visualChart.Snapshot.Style.ValueGridLineColor);
            Assert.Equal(2D, visualChart.Snapshot.Style.ValueGridLineWidth);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartGridlineStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("stroke=\"#A855F7\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-width=\"2\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(168, 85, 247),
                    tolerance: 8),
                "Expected the exported chart to include the authored gridline color at the authored width.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartGridlineDashIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("GridlineDash");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Gridline Dash");
            chart.SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: "A855F7", lineWidthPoints: 2D);
            SetFirstChartValueGridlineDash(document, A.PresetLineDashValues.Dash);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(true, visualChart.Snapshot.Style!.ShowValueGridLines);
            Assert.Equal(OfficeColor.FromRgb(168, 85, 247), visualChart.Snapshot.Style.ValueGridLineColor);
            Assert.Equal(2D, visualChart.Snapshot.Style.ValueGridLineWidth);
            Assert.Equal(OfficeStrokeDashStyle.Dash, visualChart.Snapshot.Style.ValueGridLineDashStyle);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartGridlineStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("stroke=\"#A855F7\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray=\"8 4\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(168, 85, 247),
                    tolerance: 8),
                "Expected the exported chart to include the authored dashed gridline color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesSuppressedChartGridlinesIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("NoGridlines");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "No Gridlines");
            chart.SetValueAxisGridlines(showMajor: false);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelVisualChart visualChart = Assert.Single(range.CreateVisualSnapshot(options).Charts);
            string svg = range.ToSvg(options);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.False(visualChart.Snapshot.Style!.ShowGridLines);
            Assert.DoesNotContain("#E2E8F0", svg, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesCategoryAndValueChartAxisLineColorsIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("AxisColor");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Axis Color");
            chart.SetCategoryAxisLine(lineColor: "DC2626");
            chart.SetValueAxisLine(lineColor: "2563EB");

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 3D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(OfficeColor.FromRgb(220, 38, 38), visualChart.Snapshot.Style!.CategoryAxisColor);
            Assert.Equal(OfficeColor.FromRgb(37, 99, 235), visualChart.Snapshot.Style.ValueAxisColor);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartAxisStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("#DC2626", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2563EB", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            bool hasCategoryAxisPixel = ContainsPixelNear(
                rendered!,
                visualChart.X * options.Scale,
                visualChart.Y * options.Scale,
                (visualChart.X + visualChart.Width) * options.Scale,
                (visualChart.Y + visualChart.Height) * options.Scale,
                OfficeColor.FromRgb(220, 38, 38),
                tolerance: 12)
                || ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(234, 147, 147),
                    tolerance: 38);
            bool hasValueAxisPixel = ContainsPixelNear(
                rendered!,
                visualChart.X * options.Scale,
                visualChart.Y * options.Scale,
                (visualChart.X + visualChart.Width) * options.Scale,
                (visualChart.Y + visualChart.Height) * options.Scale,
                OfficeColor.FromRgb(37, 99, 235),
                tolerance: 12)
                || ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(146, 178, 245),
                    tolerance: 38);
            Assert.True(hasCategoryAxisPixel, "Expected the exported chart to include the authored category axis line color.");
            Assert.True(hasValueAxisPixel, "Expected the exported chart to include the authored value axis line color.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesSuppressedChartAxisLinesIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("NoAxisLines");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "No Axis Lines");
            chart.SetCategoryAxisLine(noLine: true);
            chart.SetValueAxisLine(noLine: true);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelVisualChart visualChart = Assert.Single(range.CreateVisualSnapshot(options).Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);

            Assert.NotNull(visualChart.Snapshot.Layout);
            Assert.False(visualChart.Snapshot.Layout!.ShowCategoryAxisLine);
            Assert.False(visualChart.Snapshot.Layout.ShowValueAxisLine);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartAxisLineWidthIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("AxisWidth");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Axis Width");
            chart.SetValueAxisLine(lineColor: "DC2626", lineWidthPoints: 2D);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(OfficeColor.FromRgb(220, 38, 38), visualChart.Snapshot.Style!.ValueAxisColor);
            Assert.Equal(2D, visualChart.Snapshot.Style.ValueAxisLineWidth);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartAxisStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("stroke=\"#DC2626\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-width=\"2\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(220, 38, 38),
                    tolerance: 12),
                "Expected the exported chart to include the authored axis line color at the authored width.");
        }

        [Fact]
        public void ExcelRange_ImageExportCarriesChartAxisLineDashIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("AxisDash");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Actual");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 250, heightPixels: 165, type: ExcelChartType.ColumnClustered, title: "Axis Dash");
            chart.SetValueAxisLine(lineColor: "DC2626", lineWidthPoints: 2D);
            SetFirstChartValueAxisDash(document, A.PresetLineDashValues.Dot);

            ExcelRange range = sheet.Range("A1:H9");
            var options = new ExcelImageExportOptions { ShowGridlines = false, Scale = 2D };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            ExcelVisualChart visualChart = Assert.Single(snapshot.Charts);
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = range.ToSvg(options);

            Assert.NotNull(visualChart.Snapshot.Style);
            Assert.Equal(OfficeColor.FromRgb(220, 38, 38), visualChart.Snapshot.Style!.ValueAxisColor);
            Assert.Equal(2D, visualChart.Snapshot.Style.ValueAxisLineWidth);
            Assert.Equal(OfficeStrokeDashStyle.Dot, visualChart.Snapshot.Style.ValueAxisLineDashStyle);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ChartAxisStyleApproximation);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("stroke=\"#DC2626\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("stroke-dasharray=\"2 4\"", svg, StringComparison.OrdinalIgnoreCase);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(
                ContainsPixelNear(
                    rendered!,
                    visualChart.X * options.Scale,
                    visualChart.Y * options.Scale,
                    (visualChart.X + visualChart.Width) * options.Scale,
                    (visualChart.Y + visualChart.Height) * options.Scale,
                    OfficeColor.FromRgb(220, 38, 38),
                    tolerance: 12),
                "Expected the exported chart to include the authored dashed axis line color.");
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnsupportedChartTrendlines() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Trendline");
            sheet.CellValue(1, 1, "Month");
            sheet.CellValue(1, 2, "Revenue");
            sheet.CellValue(2, 1, "Jan");
            sheet.CellValue(2, 2, 120);
            sheet.CellValue(3, 1, "Feb");
            sheet.CellValue(3, 2, 180);
            sheet.CellValue(4, 1, "Mar");
            sheet.CellValue(4, 2, 160);
            ExcelChart chart = sheet.AddChartFromRange("A1:B4", row: 1, column: 4, widthPixels: 240, heightPixels: 150, type: ExcelChartType.Line, title: "Trendline");
            chart.SetSeriesTrendline(0, C.TrendlineValues.Linear, displayEquation: true, displayRSquared: true, lineColor: "FF0000", lineWidthPoints: 1);

            OfficeImageExportResult png = sheet.Range("A1:H9").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ChartTrendlineUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Trendline!" + chart.Name, diagnostic.Source);
            Assert.DoesNotContain(png.Diagnostics, item => item.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelDocument_ExportsWorkbookImagesAndSavesFolder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-Excel-Images-" + Guid.NewGuid().ToString("N"));
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet first = document.AddWorksheet("First");
            first.CellValue(1, 1, "One");
            ExcelSheet second = document.AddWorksheet("Second");
            second.CellValue(1, 1, "Two");

            IReadOnlyList<OfficeImageExportResult> results = document.ExportImages(OfficeImageExportFormat.Png);
            IReadOnlyList<OfficeImageExportResult> saved = document.SaveAsImages(folder);

            Assert.Equal(2, results.Count);
            Assert.Equal(2, saved.Count);
            Assert.True(File.Exists(Path.Combine(folder, "First.png")));
            Assert.True(File.Exists(Path.Combine(folder, "Second.png")));
            OfficeImageInfo info = OfficeImageReader.Identify(results[0].Bytes);
            Assert.Equal(OfficeImageFormat.Png, info.Format);
        }

        [Fact]
        public void ExcelDocument_ToImagesFluentExportsSelectedSheetsThroughSharedBatchBuilder() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-Excel-Fluent-Images-" + Guid.NewGuid().ToString("N"));
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet first = document.AddWorksheet("First");
            first.CellValue(1, 1, "One");
            ExcelSheet second = document.AddWorksheet("Second");
            second.CellValue(1, 1, "Two");

            IReadOnlyList<OfficeImageExportResult> results = document
                .ToImages()
                .ForSheets("Second")
                .WithoutGridlines()
                .As(OfficeImageExportFormat.Svg)
                .Save(folder);

            OfficeImageExportResult result = Assert.Single(results);
            Assert.Equal(OfficeImageExportFormat.Svg, result.Format);
            Assert.Equal("Second", result.Name);
            Assert.True(File.Exists(Path.Combine(folder, "Second.svg")));
            Assert.False(File.Exists(Path.Combine(folder, "First.svg")));
            string svg = System.Text.Encoding.UTF8.GetString(result.Bytes);
            Assert.Contains("Two", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("One", svg, StringComparison.Ordinal);
            Assert.Empty(result.Diagnostics);
        }

        [Fact]
        public void ExcelDocument_ToImagesSnapshotsCallerOwnedSheetNames() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            document.AddWorksheet("First").CellValue(1, 1, "One");
            document.AddWorksheet("Second").CellValue(1, 1, "Two");
            var selectedSheets = new List<string> { "First" };
            var options = new ExcelWorkbookImageExportOptions {
                SheetNames = selectedSheets
            };
            ExcelWorkbookImageExportBuilder builder = document.ToImages(options);

            selectedSheets[0] = "Second";
            IReadOnlyList<OfficeImageExportResult> results = builder
                .As(OfficeImageExportFormat.Svg)
                .Export();

            OfficeImageExportResult result = Assert.Single(results);
            Assert.Equal("First", result.Name);
            Assert.Contains(
                "One",
                System.Text.Encoding.UTF8.GetString(result.Bytes),
                StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelDocument_ImageExportRejectsMissingRequestedSheetNames() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            string folder = Path.Combine(Path.GetTempPath(), "OfficeIMO-Excel-Missing-Sheet-" + Guid.NewGuid().ToString("N"));
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet first = document.AddWorksheet("First");
            first.CellValue(1, 1, "One");
            ExcelSheet second = document.AddWorksheet("Second");
            second.CellValue(1, 1, "Two");

            var options = new ExcelWorkbookImageExportOptions { SheetNames = new[] { "Second", "Missing" } };
            ArgumentException exportException = Assert.Throws<ArgumentException>(() => document.ExportImages(OfficeImageExportFormat.Png, options));
            ArgumentException saveException = Assert.Throws<ArgumentException>(() => document.ToImages().ForSheets("Missing").Save(folder));

            Assert.Contains("Missing", exportException.Message, StringComparison.Ordinal);
            Assert.Contains("Missing", saveException.Message, StringComparison.Ordinal);
            Assert.Empty(Directory.Exists(folder) ? Directory.GetFiles(folder) : Array.Empty<string>());
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsCellVerticalTextAlignment() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Align");
            sheet.CellValue(1, 1, "Top");
            sheet.CellValue(1, 2, "Mid");
            sheet.CellValue(1, 3, "Bot");
            sheet.SetRowHeight(1, 60);
            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 14);
            sheet.SetColumnWidth(3, 14);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Top);
            sheet.CellVerticalAlign(1, 2, VerticalAlignmentValues.Center);
            sheet.CellVerticalAlign(1, 3, VerticalAlignmentValues.Bottom);

            ExcelRange range = sheet.Range("A1:C1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            int topY = MinDarkPixelY(rendered!, snapshot.Cells[0]);
            int middleY = MinDarkPixelY(rendered!, snapshot.Cells[1]);
            int bottomY = MinDarkPixelY(rendered!, snapshot.Cells[2]);
            Assert.True(topY < middleY, $"Expected top-aligned text to start above middle text. top={topY}, middle={middleY}");
            Assert.True(middleY < bottomY, $"Expected middle-aligned text to start above bottom text. middle={middleY}, bottom={bottomY}");
        }

        [Fact]
        public void ExcelRange_ImageExportDefaultsToBottomVerticalTextAlignment() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("DefaultAlign");
            sheet.CellValue(1, 1, "Top");
            sheet.CellValue(1, 2, "Default");
            sheet.CellValue(1, 3, "Bottom");
            sheet.SetRowHeight(1, 60);
            sheet.SetColumnWidth(1, 14);
            sheet.SetColumnWidth(2, 14);
            sheet.SetColumnWidth(3, 14);
            sheet.CellVerticalAlign(1, 1, VerticalAlignmentValues.Top);
            sheet.CellVerticalAlign(1, 3, VerticalAlignmentValues.Bottom);

            ExcelRange range = sheet.Range("A1:C1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            int topY = MinDarkPixelY(rendered!, snapshot.Cells[0]);
            int defaultY = MinDarkPixelY(rendered!, snapshot.Cells[1]);
            int bottomY = MinDarkPixelY(rendered!, snapshot.Cells[2]);
            Assert.True(defaultY > topY, $"Expected default text to start below top-aligned text. top={topY}, default={defaultY}");
            Assert.True(Math.Abs(defaultY - bottomY) <= 3, $"Expected default text to align like bottom text. default={defaultY}, bottom={bottomY}");
        }

        [Fact]
        public void ExcelRange_ImageExportWrapsCellTextAcrossPngAndSvg() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Wrap");
            sheet.CellValue(1, 1, "Alpha Beta Gamma Delta");
            sheet.SetColumnWidth(1, 8);
            sheet.SetRowHeight(1, 54);
            sheet.WrapCells(1, 1, 1);

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            (int MinY, int MaxY) extent = DarkPixelYExtent(rendered!, snapshot.Cells[0]);
            Assert.True(extent.MaxY - extent.MinY > 20, $"Expected wrapped text to occupy multiple visible lines, got y extent {extent.MinY}-{extent.MaxY}.");
            Assert.Contains("<clipPath", svg, StringComparison.Ordinal);
            Assert.Contains("Alpha", svg, StringComparison.Ordinal);
            Assert.Contains("Gamma", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsClippedCellText() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Clip");
            sheet.CellValue(1, 1, "This text is intentionally much too long for the rendered cell");
            sheet.SetColumnWidth(1, 6);

            OfficeImageExportResult png = sheet.Range("A1:A1").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextClipped);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Clip!A1", diagnostic.Source);
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsFontSizeAcrossPngAndSvg() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("FontSize");
            sheet.CellValue(1, 1, "Small");
            sheet.CellValue(2, 1, "Large");
            sheet.CellAt(1, 1).SetFontSize(8);
            sheet.CellAt(2, 1).SetFontSize(18);
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 32);
            sheet.SetRowHeight(2, 42);

            ExcelRange range = sheet.Range("A1:A2");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(8D, snapshot.Cells[0].Style.FontSize);
            Assert.Equal(18D, snapshot.Cells[1].Style.FontSize);
            Assert.Contains("font-size=\"8\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-size=\"18\"", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            (int SmallMinY, int SmallMaxY) smallExtent = DarkPixelYExtent(rendered!, snapshot.Cells[0]);
            (int LargeMinY, int LargeMaxY) largeExtent = DarkPixelYExtent(rendered!, snapshot.Cells[1]);
            Assert.True(
                largeExtent.LargeMaxY - largeExtent.LargeMinY > smallExtent.SmallMaxY - smallExtent.SmallMinY,
                $"Expected larger font to occupy more vertical pixels. small={smallExtent.SmallMinY}-{smallExtent.SmallMaxY}, large={largeExtent.LargeMinY}-{largeExtent.LargeMaxY}");
        }

        [Fact]
        public void ExcelRange_ImageExportShrinksCellTextToFitWithoutClippingDiagnostic() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Shrink");
            sheet.CellValue(1, 1, "Shrink To Fit");
            sheet.SetColumnWidth(1, 7);
            sheet.CellAt(1, 1).SetFontSize(14).SetShrinkToFit();

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.True(snapshot.Cells[0].Style.ShrinkToFit);
            Assert.Equal(14D, snapshot.Cells[0].Style.FontSize);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped);
            Assert.Contains("Shrink To Fit", svg, StringComparison.Ordinal);
            double renderedFontSize = ExtractFirstSvgFontSize(svg);
            Assert.True(renderedFontSize < 14D, "Expected shrink-to-fit SVG text to use a smaller font size than the source style.");
            Assert.True(renderedFontSize > 1D, "Expected shrink-to-fit SVG text to stay visible.");
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            DarkPixelYExtent(rendered!, snapshot.Cells[0]);
        }

        [Fact]
        public void ExcelRange_ImageExportHonorsTextRotationAcrossPngAndSvg() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Rotate");
            sheet.CellValue(1, 1, "Rotated Header");
            sheet.SetColumnWidth(1, 9);
            sheet.SetRowHeight(1, 86);
            sheet.CellAt(1, 1)
                .SetTextRotation(90)
                .SetFontColor("010203")
                .SetBold()
                .SetItalic()
                .SetUnderline();

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(90, snapshot.Cells[0].Style.TextRotation);
            Assert.Contains("rotate(-90", svg, StringComparison.Ordinal);
            Assert.Contains("text-anchor=\"middle\"", svg, StringComparison.Ordinal);
            Assert.Contains("fill=\"#010203\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline\"", svg, StringComparison.Ordinal);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextRotationApproximation);
            Assert.Equal("Rotate!A1", diagnostic.Source);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            (int MinY, int MaxY) yExtent = DarkPixelYExtent(rendered!, snapshot.Cells[0]);
            (int MinX, int MaxX) xExtent = DarkPixelXExtent(rendered!, snapshot.Cells[0]);
            Assert.True(
                yExtent.MaxY - yExtent.MinY > xExtent.MaxX - xExtent.MinX,
                $"Expected 90 degree text to be taller than wide. x={xExtent.MinX}-{xExtent.MaxX}, y={yExtent.MinY}-{yExtent.MaxY}");
        }

        [Fact]
        public void ExcelRange_ImageExportClipsRotatedPngTextToCellBounds() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("ClipRotate");
            sheet.CellValue(1, 1, "Rotated Header");
            sheet.SetColumnWidth(1, 8);
            sheet.SetColumnWidth(2, 12);
            sheet.SetRowHeight(1, 52);
            sheet.CellAt(1, 1).SetTextRotation(45);

            ExcelRange range = sheet.Range("A1:B1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.True(ContainsDarkPixel(rendered!, snapshot.Cells[0]));
            Assert.False(ContainsDarkPixel(rendered!, snapshot.Cells[1]));
        }

        [Fact]
        public void ExcelRange_ImageExportRendersStackedTextRotationAcrossPngAndSvg() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Stacked");
            sheet.CellValue(1, 1, "Stacked");
            sheet.SetColumnWidth(1, 5);
            sheet.SetRowHeight(1, 96);
            sheet.CellAt(1, 1)
                .SetTextRotation(255)
                .SetFontColor("010203")
                .SetBold();

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = range.ToSvg(new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(255, snapshot.Cells[0].Style.TextRotation);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellStackedTextRotationUnsupported);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellTextRotationApproximation);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Stacked!A1", diagnostic.Source);
            Assert.DoesNotContain("rotate(", svg, StringComparison.Ordinal);
            Assert.Contains("fill=\"#010203\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.True(CountOccurrences(svg, "<text") >= "Stacked".Length, "Expected stacked SVG output to emit one visible text element per stacked character.");
            Assert.Contains(">S</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">t</text>", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            (int MinY, int MaxY) yExtent = DarkPixelYExtent(rendered!, snapshot.Cells[0]);
            (int MinX, int MaxX) xExtent = DarkPixelXExtent(rendered!, snapshot.Cells[0]);
            Assert.True(
                yExtent.MaxY - yExtent.MinY > xExtent.MaxX - xExtent.MinX,
                $"Expected stacked text to be taller than wide. x={xExtent.MinX}-{xExtent.MaxX}, y={yExtent.MinY}-{yExtent.MaxY}");
        }

        private static byte[] CreateSolidPng(int width, int height, OfficeColor color) {
            OfficeRasterImage image = new OfficeRasterImage(width, height, OfficeColor.Transparent);
            image.Fill(color);
            return OfficePngWriter.Encode(image);
        }

        private static byte[] CreateSolidOpaquePng(int width, int height, OfficeColor color) {
            byte[] scanlines = new byte[checked(height * (1 + width * 3))];
            int offset = 0;
            for (int row = 0; row < height; row++) {
                scanlines[offset++] = 0;
                for (int column = 0; column < width; column++) {
                    scanlines[offset++] = color.R;
                    scanlines[offset++] = color.G;
                    scanlines[offset++] = color.B;
                }
            }

            return OfficePngWriter.EncodeScanlines(width, height, bitDepth: 8, colorType: 2, scanlines, OfficePngCompression.Stored);
        }

        private static byte[] CreateSinglePixelGif() =>
            Convert.FromBase64String("R0lGODlhAQABAIAAAAAAAP///ywAAAAAAQABAAACAUwAOw==");

        private static byte[] CreateTruncatedPngHeader(int width, int height) {
            byte[] bytes = new byte[33];
            bytes[0] = 137;
            bytes[1] = 80;
            bytes[2] = 78;
            bytes[3] = 71;
            bytes[4] = 13;
            bytes[5] = 10;
            bytes[6] = 26;
            bytes[7] = 10;
            bytes[8] = 0;
            bytes[9] = 0;
            bytes[10] = 0;
            bytes[11] = 13;
            bytes[12] = (byte)'I';
            bytes[13] = (byte)'H';
            bytes[14] = (byte)'D';
            bytes[15] = (byte)'R';
            WriteInt32BigEndian(bytes, 16, width);
            WriteInt32BigEndian(bytes, 20, height);
            bytes[24] = 8;
            bytes[25] = 6;
            return bytes;
        }

        private static byte[] CreateBmp24(int width, int height, IReadOnlyList<OfficeColor> pixels, bool topDown = false) {
            int rowStride = ((width * 24) + 31) / 32 * 4;
            int pixelOffset = 54;
            byte[] bytes = new byte[pixelOffset + (rowStride * height)];
            bytes[0] = (byte)'B';
            bytes[1] = (byte)'M';
            WriteInt32LittleEndian(bytes, 2, bytes.Length);
            WriteInt32LittleEndian(bytes, 10, pixelOffset);
            WriteInt32LittleEndian(bytes, 14, 40);
            WriteInt32LittleEndian(bytes, 18, width);
            WriteInt32LittleEndian(bytes, 22, topDown ? -height : height);
            WriteUInt16LittleEndian(bytes, 26, 1);
            WriteUInt16LittleEndian(bytes, 28, 24);

            for (int y = 0; y < height; y++) {
                int sourceY = topDown ? y : height - 1 - y;
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

        private static byte[] CreateBmp32(int width, int height, IReadOnlyList<OfficeColor> pixels) {
            int rowStride = width * 4;
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
            WriteUInt16LittleEndian(bytes, 28, 32);

            for (int y = 0; y < height; y++) {
                int sourceY = height - 1 - y;
                int rowOffset = pixelOffset + (y * rowStride);
                for (int x = 0; x < width; x++) {
                    OfficeColor color = pixels[(sourceY * width) + x];
                    int offset = rowOffset + (x * 4);
                    bytes[offset] = color.B;
                    bytes[offset + 1] = color.G;
                    bytes[offset + 2] = color.R;
                    bytes[offset + 3] = color.A;
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

        private static void WriteInt32BigEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)(value >> 24);
            bytes[offset + 1] = (byte)(value >> 16);
            bytes[offset + 2] = (byte)(value >> 8);
            bytes[offset + 3] = (byte)value;
        }

        private static void WriteUInt16LittleEndian(byte[] bytes, int offset, int value) {
            bytes[offset] = (byte)value;
            bytes[offset + 1] = (byte)(value >> 8);
        }

        private static byte[] CreateHorizontalCropPng() {
            OfficeRasterImage image = new OfficeRasterImage(80, 24, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 20, 24, OfficeColor.FromRgb(220, 38, 38));
            canvas.FillRectangle(20, 0, 40, 24, OfficeColor.FromRgb(37, 99, 235));
            canvas.FillRectangle(60, 0, 20, 24, OfficeColor.FromRgb(22, 163, 74));
            return OfficePngWriter.Encode(image);
        }

        private static byte[] CreateRotationProbePng() {
            OfficeRasterImage image = new OfficeRasterImage(120, 54, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 120, 54, OfficeColor.FromRgb(37, 99, 235));
            canvas.FillRectangle(0, 0, 38, 54, OfficeColor.FromRgb(20, 184, 166));
            canvas.FillRectangle(82, 0, 38, 54, OfficeColor.FromRgb(124, 58, 237));
            canvas.DrawRectangle(2, 2, 116, 50, OfficeColor.White, 3);
            canvas.DrawLine(46, 18, 74, 18, OfficeColor.White, 4);
            canvas.DrawLine(38, 34, 82, 34, OfficeColor.White, 4);
            return OfficePngWriter.Encode(image);
        }

        private static byte[] CreateTransformProbePng() {
            OfficeRasterImage image = new OfficeRasterImage(160, 60, OfficeColor.Transparent);
            OfficeRasterCanvas canvas = new OfficeRasterCanvas(image);
            canvas.FillRectangle(0, 0, 40, 60, OfficeColor.FromRgb(220, 38, 38));
            canvas.FillRectangle(40, 0, 70, 60, OfficeColor.FromRgb(37, 99, 235));
            canvas.FillRectangle(110, 0, 50, 60, OfficeColor.FromRgb(22, 163, 74));
            canvas.DrawRectangle(42, 5, 66, 50, OfficeColor.White, 3);
            canvas.DrawLine(60, 22, 96, 22, OfficeColor.White, 4);
            canvas.DrawLine(54, 38, 102, 38, OfficeColor.White, 4);
            return OfficePngWriter.Encode(image);
        }

        private static void SetFirstChartValueGridlineDash(ExcelDocument document, A.PresetLineDashValues dashStyle) {
            ChartPart chartPart = GetFirstChartPart(document);
            C.ValueAxis valueAxis = chartPart.ChartSpace.Descendants<C.ValueAxis>().First();
            C.MajorGridlines majorGridlines = valueAxis.GetFirstChild<C.MajorGridlines>() ?? new C.MajorGridlines();
            if (majorGridlines.Parent == null) {
                valueAxis.Append(majorGridlines);
            }

            C.ChartShapeProperties properties = majorGridlines.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (properties.Parent == null) {
                majorGridlines.Append(properties);
            }

            SetPresetDash(properties, dashStyle);
            chartPart.ChartSpace.Save();
        }

        private static void SetFirstChartValueAxisDash(ExcelDocument document, A.PresetLineDashValues dashStyle) {
            ChartPart chartPart = GetFirstChartPart(document);
            C.ValueAxis valueAxis = chartPart.ChartSpace.Descendants<C.ValueAxis>().First();
            C.ShapeProperties properties = valueAxis.GetFirstChild<C.ShapeProperties>() ?? new C.ShapeProperties();
            if (properties.Parent == null) {
                valueAxis.Append(properties);
            }

            SetPresetDash(properties, dashStyle);
            chartPart.ChartSpace.Save();
        }

        private static void SetFirstChartSeriesDash(ExcelDocument document, A.PresetLineDashValues dashStyle) {
            ChartPart chartPart = GetFirstChartPart(document);
            C.LineChartSeries series = chartPart.ChartSpace.Descendants<C.LineChartSeries>().First();
            C.ChartShapeProperties properties = series.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (properties.Parent == null) {
                series.Append(properties);
            }

            SetPresetDash(properties, dashStyle);
            chartPart.ChartSpace.Save();
        }

        private static void SetFirstChartSeriesNoLine(ExcelDocument document) {
            ChartPart chartPart = GetFirstChartPart(document);
            C.LineChartSeries series = chartPart.ChartSpace.Descendants<C.LineChartSeries>().First();
            C.ChartShapeProperties properties = series.GetFirstChild<C.ChartShapeProperties>() ?? new C.ChartShapeProperties();
            if (properties.Parent == null) {
                series.Append(properties);
            }

            A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren();
            outline.Append(new A.NoFill());
            if (outline.Parent == null) {
                properties.Append(outline);
            }

            chartPart.ChartSpace.Save();
        }

        private static void SetFirstChartSeriesIndex(ExcelDocument document, uint index) {
            ChartPart chartPart = GetFirstChartPart(document);
            OpenXmlCompositeElement series = chartPart.ChartSpace.Descendants<C.LineChartSeries>().Cast<OpenXmlCompositeElement>().First();
            C.Index indexElement = series.GetFirstChild<C.Index>() ?? new C.Index();
            indexElement.Val = index;
            if (indexElement.Parent == null) {
                series.InsertAt(indexElement, 0);
            }

            C.Order orderElement = series.GetFirstChild<C.Order>() ?? new C.Order();
            orderElement.Val = index;
            if (orderElement.Parent == null) {
                series.InsertAfter(orderElement, indexElement);
            }

            chartPart.ChartSpace.Save();
        }

        private static void SetFirstDataBarThresholdFormula(ExcelSheet sheet, string formula) {
            DataBar dataBar = sheet.WorksheetPart.Worksheet.Descendants<DataBar>().First();
            ConditionalFormatValueObject threshold = dataBar.Elements<ConditionalFormatValueObject>().First();
            threshold.Type = ConditionalFormatValueObjectValues.Formula;
            threshold.Val = formula;
            sheet.WorksheetPart.Worksheet.Save();
        }

        private static void SetFirstColorScaleThresholdFormula(ExcelSheet sheet, string formula) {
            ColorScale colorScale = sheet.WorksheetPart.Worksheet.Descendants<ColorScale>().First();
            ConditionalFormatValueObject threshold = colorScale.Elements<ConditionalFormatValueObject>().First();
            threshold.Type = ConditionalFormatValueObjectValues.Formula;
            threshold.Val = formula;
            sheet.WorksheetPart.Worksheet.Save();
        }

        private static void SetFirstChartAreaNoFill(ExcelDocument document) {
            ChartPart chartPart = GetFirstChartPart(document);
            C.ShapeProperties properties = chartPart.ChartSpace.GetFirstChild<C.ShapeProperties>() ?? new C.ShapeProperties();
            if (properties.Parent == null) {
                chartPart.ChartSpace.Append(properties);
            }

            properties.RemoveAllChildren<A.SolidFill>();
            properties.RemoveAllChildren<A.NoFill>();
            properties.PrependChild(new A.NoFill());
            A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren();
            outline.Append(new A.NoFill());
            if (outline.Parent == null) {
                properties.Append(outline);
            }

            chartPart.ChartSpace.Save();
        }

        private static void MoveFirstChartToAbsoluteAnchor(string filePath, int xPixels, int yPixels, int widthPixels, int heightPixels) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? throw new InvalidOperationException("Worksheet has no drawings part.");
            Xdr.WorksheetDrawing worksheetDrawing = drawingsPart.WorksheetDrawing ?? throw new InvalidOperationException("Worksheet has no drawing.");
            OpenXmlCompositeElement anchor = worksheetDrawing.ChildElements
                .OfType<OpenXmlCompositeElement>()
                .First(element => element.GetFirstChild<Xdr.GraphicFrame>() != null);
            Xdr.GraphicFrame frame = (Xdr.GraphicFrame)anchor.GetFirstChild<Xdr.GraphicFrame>()!.CloneNode(true);
            Xdr.ClientData clientData = (Xdr.ClientData?)anchor.GetFirstChild<Xdr.ClientData>()?.CloneNode(true) ?? new Xdr.ClientData();
            var absoluteAnchor = new Xdr.AbsoluteAnchor(
                new Xdr.Position { X = xPixels * 9525L, Y = yPixels * 9525L },
                new Xdr.Extent { Cx = widthPixels * 9525L, Cy = heightPixels * 9525L },
                frame,
                clientData);
            worksheetDrawing.ReplaceChild(absoluteAnchor, anchor);
            worksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void SetFirstChartAreaDash(ExcelDocument document, A.PresetLineDashValues dashStyle) {
            ChartPart chartPart = GetFirstChartPart(document);
            C.ShapeProperties properties = chartPart.ChartSpace.GetFirstChild<C.ShapeProperties>() ?? new C.ShapeProperties();
            if (properties.Parent == null) {
                chartPart.ChartSpace.Append(properties);
            }

            SetPresetDash(properties, dashStyle);
            chartPart.ChartSpace.Save();
        }

        private static void SetFirstChartPlotAreaDash(ExcelDocument document, A.PresetLineDashValues dashStyle) {
            ChartPart chartPart = GetFirstChartPart(document);
            C.PlotArea plotArea = chartPart.ChartSpace.Descendants<C.PlotArea>().First();
            C.ShapeProperties properties = plotArea.GetFirstChild<C.ShapeProperties>() ?? new C.ShapeProperties();
            if (properties.Parent == null) {
                plotArea.Append(properties);
            }

            SetPresetDash(properties, dashStyle);
            chartPart.ChartSpace.Save();
        }

        private static void AddFirstChartTitleTextEffect(ExcelDocument document) {
            ChartPart chartPart = GetFirstChartPart(document);
            A.RunProperties runProperties = chartPart.ChartSpace.Descendants<A.RunProperties>().First();
            runProperties.Append(new A.EffectList());
            chartPart.ChartSpace.Save();
        }

        private static ChartPart GetFirstChartPart(ExcelDocument document) =>
            document.WorkbookPartRoot
                .WorksheetParts
                .Select(part => part.DrawingsPart)
                .Where(part => part != null)
                .SelectMany(part => part!.ChartParts)
                .First();

        private static void SetPresetDash(OpenXmlCompositeElement properties, A.PresetLineDashValues dashStyle) {
            A.Outline outline = properties.GetFirstChild<A.Outline>() ?? new A.Outline();
            outline.RemoveAllChildren<A.PresetDash>();
            outline.Append(new A.PresetDash { Val = dashStyle });
            if (outline.Parent == null) {
                properties.Append(outline);
            }
        }

        private static void AddMalformedTimePeriodRuleWithFill(ExcelSheet sheet, string range) {
            sheet.AddConditionalTimePeriodRule(range, TimePeriodValues.Today, fillColor: "C6EFCE");
            ConditionalFormattingRule rule = sheet.WorksheetPart.Worksheet!
                .Elements<ConditionalFormatting>()
                .Where(conditional => string.Equals(conditional.SequenceOfReferences?.InnerText, range, StringComparison.Ordinal))
                .SelectMany(conditional => conditional.Elements<ConditionalFormattingRule>())
                .First(item => item.Type?.Value == ConditionalFormatValues.TimePeriod);
            rule.TimePeriod = null;
            sheet.WorksheetPart.Worksheet.Save();
        }

        private static void MarkAboveAverageRuleAsStdDev(ExcelSheet sheet, string range, int stdDev) {
            ConditionalFormattingRule rule = sheet.WorksheetPart.Worksheet!
                .Elements<ConditionalFormatting>()
                .Where(conditional => string.Equals(conditional.SequenceOfReferences?.InnerText, range, StringComparison.Ordinal))
                .SelectMany(conditional => conditional.Elements<ConditionalFormattingRule>())
                .First(item => item.Type?.Value == ConditionalFormatValues.AboveAverage);
            rule.StdDev = stdDev;
            sheet.WorksheetPart.Worksheet.Save();
        }

        private static void AddTwoCellAnchoredImage(
            string filePath,
            byte[] imageBytes,
            string fromColumnOffset = "0",
            string toColumnOffset = "0",
            string toColumnId = "3") {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (MemoryStream stream = new MemoryStream(imageBytes)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = drawingsPart.GetIdOfPart(imagePart);
            drawingsPart.WorksheetDrawing.Append(new Xdr.TwoCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("0"),
                    new Xdr.ColumnOffset(fromColumnOffset),
                    new Xdr.RowId("0"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId(toColumnId),
                    new Xdr.ColumnOffset(toColumnOffset),
                    new Xdr.RowId("2"),
                    new Xdr.RowOffset("0")),
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 77U, Name = "TwoCellBanner" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddAbsoluteAnchoredImage(string filePath, byte[] imageBytes, int xPixels, int yPixels, int widthPixels, int heightPixels) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (MemoryStream stream = new MemoryStream(imageBytes)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = drawingsPart.GetIdOfPart(imagePart);
            drawingsPart.WorksheetDrawing.Append(new Xdr.AbsoluteAnchor(
                new Xdr.Position { X = xPixels * 9525L, Y = yPixels * 9525L },
                new Xdr.Extent { Cx = widthPixels * 9525L, Cy = heightPixels * 9525L },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 81U, Name = "AbsoluteBanner" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddCroppedImage(string filePath, byte[] imageBytes) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (MemoryStream stream = new MemoryStream(imageBytes)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = drawingsPart.GetIdOfPart(imagePart);
            drawingsPart.WorksheetDrawing.Append(new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("0"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("0"),
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = 96L * 9525L, Cy = 30L * 9525L },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 78U, Name = "CroppedBanner" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.SourceRectangle { Left = 25000, Right = 25000 },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddRotatedImage(string filePath, byte[] imageBytes) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (MemoryStream stream = new MemoryStream(imageBytes)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = drawingsPart.GetIdOfPart(imagePart);
            drawingsPart.WorksheetDrawing.Append(new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("1"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("1"),
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = 120L * 9525L, Cy = 54L * 9525L },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 79U, Name = "RotatedBanner" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 120L * 9525L, Cy = 54L * 9525L }) { Rotation = 30 * 60000 },
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void AddTransformedImage(string filePath, byte[] imageBytes) {
            using SpreadsheetDocument spreadsheet = SpreadsheetDocument.Open(filePath, true);
            WorksheetPart worksheetPart = spreadsheet.WorkbookPart!.WorksheetParts.First();
            DrawingsPart drawingsPart = worksheetPart.DrawingsPart ?? worksheetPart.AddNewPart<DrawingsPart>();
            drawingsPart.WorksheetDrawing ??= new Xdr.WorksheetDrawing();

            if (worksheetPart.Worksheet!.Elements<X.Drawing>().FirstOrDefault() == null) {
                worksheetPart.Worksheet.Append(new X.Drawing { Id = worksheetPart.GetIdOfPart(drawingsPart) });
            }

            ImagePart imagePart = drawingsPart.AddImagePart(ImagePartType.Png);
            using (MemoryStream stream = new MemoryStream(imageBytes)) {
                imagePart.FeedData(stream);
            }

            string relationshipId = drawingsPart.GetIdOfPart(imagePart);
            drawingsPart.WorksheetDrawing.Append(new Xdr.OneCellAnchor(
                new Xdr.FromMarker(
                    new Xdr.ColumnId("1"),
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("2"),
                    new Xdr.RowOffset("0")),
                new Xdr.Extent { Cx = 132L * 9525L, Cy = 56L * 9525L },
                new Xdr.Picture(
                    new Xdr.NonVisualPictureProperties(
                        new Xdr.NonVisualDrawingProperties { Id = 80U, Name = "TransformedBanner" },
                        new Xdr.NonVisualPictureDrawingProperties(new A.PictureLocks { NoChangeAspect = true })),
                    new Xdr.BlipFill(
                        new A.Blip { Embed = relationshipId },
                        new A.SourceRectangle { Left = 25000 },
                        new A.Stretch(new A.FillRectangle())),
                    new Xdr.ShapeProperties(
                        new A.Transform2D(
                            new A.Offset { X = 0L, Y = 0L },
                            new A.Extents { Cx = 132L * 9525L, Cy = 56L * 9525L }) {
                            Rotation = 30 * 60000,
                            HorizontalFlip = true
                        },
                        new A.PresetGeometry(new A.AdjustValueList()) { Preset = A.ShapeTypeValues.Rectangle })),
                new Xdr.ClientData()));
            drawingsPart.WorksheetDrawing.Save();
            worksheetPart.Worksheet.Save();
        }

        private static void ApplyThemeBackedCellStyle(ExcelDocument document, ExcelSheet sheet) {
            document.EnsureWorkbookThemeAndStyles();
            WorkbookPart workbookPart = document.WorkbookPartRoot;
            Stylesheet stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;

            Fonts fonts = stylesheet.Fonts ??= new Fonts();
            uint fontId = (uint)fonts.Elements<Font>().Count();
            fonts.Append(new Font(
                new FontSize { Val = 11D },
                new FontName { Val = "Calibri" },
                new Color { Theme = 5U, Tint = -0.5D }));
            fonts.Count = (uint)fonts.Elements<Font>().Count();

            Fills fills = stylesheet.Fills ??= new Fills();
            uint fillId = (uint)fills.Elements<Fill>().Count();
            fills.Append(new Fill(new PatternFill {
                PatternType = PatternValues.Solid,
                ForegroundColor = new ForegroundColor { Theme = 4U, Tint = 0.4D },
                BackgroundColor = new BackgroundColor { Indexed = 64U }
            }));
            fills.Count = (uint)fills.Elements<Fill>().Count();

            Borders borders = stylesheet.Borders ??= new Borders();
            uint borderId = (uint)borders.Elements<Border>().Count();
            borders.Append(new Border(
                new LeftBorder(new Color { Theme = 6U }) { Style = BorderStyleValues.Thin },
                new RightBorder(new Color { Theme = 6U }) { Style = BorderStyleValues.Thin },
                new TopBorder(new Color { Theme = 6U }) { Style = BorderStyleValues.Thin },
                new BottomBorder(new Color { Theme = 6U }) { Style = BorderStyleValues.Thin },
                new DiagonalBorder()));
            borders.Count = (uint)borders.Elements<Border>().Count();

            CellFormats formats = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            uint styleIndex = (uint)formats.Elements<CellFormat>().Count();
            formats.Append(new CellFormat {
                FontId = fontId,
                FillId = fillId,
                BorderId = borderId,
                FormatId = 0U,
                ApplyFont = true,
                ApplyFill = true,
                ApplyBorder = true
            });
            formats.Count = (uint)formats.Elements<CellFormat>().Count();

            Worksheet worksheet = sheet.WorksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            Cell cell = worksheet.Descendants<Cell>().First(item => item.CellReference?.Value == "A1");
            cell.StyleIndex = styleIndex;
            stylesheet.Save();
        }

        private static void ApplyCellTextIndentStyle(ExcelDocument document, ExcelSheet sheet, string cellReference, uint indent) {
            document.EnsureWorkbookThemeAndStyles();
            WorkbookPart workbookPart = document.WorkbookPartRoot;
            Stylesheet stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;
            CellFormats formats = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            uint styleIndex = (uint)formats.Elements<CellFormat>().Count();
            formats.Append(new CellFormat {
                NumberFormatId = 0U,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyAlignment = true,
                Alignment = new Alignment {
                    Horizontal = HorizontalAlignmentValues.Left,
                    Indent = indent
                }
            });
            formats.Count = (uint)formats.Elements<CellFormat>().Count();

            Worksheet worksheet = sheet.WorksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            Cell cell = worksheet.Descendants<Cell>().First(item => item.CellReference?.Value == cellReference);
            cell.StyleIndex = styleIndex;
            stylesheet.Save();
        }

        private static void ApplyBuiltInNumberFormatId(ExcelDocument document, ExcelSheet sheet, string cellReference, uint numberFormatId) {
            document.EnsureWorkbookThemeAndStyles();
            WorkbookPart workbookPart = document.WorkbookPartRoot;
            Stylesheet stylesheet = workbookPart.WorkbookStylesPart!.Stylesheet!;
            CellFormats formats = stylesheet.CellFormats ??= new CellFormats(new CellFormat());
            uint styleIndex = (uint)formats.Elements<CellFormat>().Count();
            formats.Append(new CellFormat {
                NumberFormatId = numberFormatId,
                FontId = 0U,
                FillId = 0U,
                BorderId = 0U,
                FormatId = 0U,
                ApplyNumberFormat = true
            });
            formats.Count = (uint)formats.Elements<CellFormat>().Count();

            Worksheet worksheet = sheet.WorksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            Cell cell = worksheet.Descendants<Cell>().First(item => item.CellReference?.Value == cellReference);
            cell.StyleIndex = styleIndex;
            stylesheet.Save();
        }

        private static void AssertPixelNear(OfficeRasterImage image, int x, int y, OfficeColor expected, int tolerance) {
            OfficeColor actual = image.GetPixel(x, y);
            Assert.True(
                Math.Abs(actual.R - expected.R) <= tolerance
                && Math.Abs(actual.G - expected.G) <= tolerance
                && Math.Abs(actual.B - expected.B) <= tolerance,
                $"Expected pixel {x},{y} near {expected}, got {actual}.");
        }

        private static void AssertCellContainsPixelNear(OfficeRasterImage image, ExcelVisualCell cell, OfficeColor expected, int tolerance) {
            double inset = 3D;
            bool hasExpectedPixel = ContainsPixelNear(
                image,
                cell.X + inset,
                cell.Y + inset,
                cell.X + cell.Width - inset,
                cell.Y + cell.Height - inset,
                expected,
                tolerance);
            Assert.True(hasExpectedPixel, $"Expected R{cell.Row}C{cell.Column} interior to contain a pixel near {expected}.");
        }

        private static int CountGreenIconPixels(OfficeRasterImage image, ExcelVisualConditionalIcon icon) {
            int left = Math.Max(0, (int)Math.Floor(icon.X));
            int top = Math.Max(0, (int)Math.Floor(icon.Y));
            int right = Math.Min(image.Width, (int)Math.Ceiling(icon.X + icon.Width));
            int bottom = Math.Min(image.Height, (int)Math.Ceiling(icon.Y + icon.Height));
            int count = 0;
            for (int y = top; y < bottom; y++) {
                for (int x = left; x < right; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.G > 120 && pixel.R < 80 && pixel.B < 120) {
                        count++;
                    }
                }
            }

            return count;
        }

        private static byte[] CreateMinimalJpegHeader() {
            return new byte[] {
                0xFF, 0xD8,
                0xFF, 0xE0, 0x00, 0x10,
                0x4A, 0x46, 0x49, 0x46, 0x00,
                0x01, 0x01, 0x00,
                0x00, 0x01, 0x00, 0x01,
                0x00, 0x00,
                0xFF, 0xC0, 0x00, 0x11,
                0x08,
                0x00, 0x01,
                0x00, 0x01,
                0x03,
                0x01, 0x11, 0x00,
                0x02, 0x11, 0x00,
                0x03, 0x11, 0x00,
                0xFF, 0xD9
            };
        }

        private static int MinDarkPixelY(OfficeRasterImage image, ExcelVisualCell cell) {
            (int MinY, _) = DarkPixelYExtent(image, cell);
            return MinY;
        }

        private static (int MinY, int MaxY) DarkPixelYExtent(OfficeRasterImage image, ExcelVisualCell cell) {
            int left = Math.Max(0, (int)Math.Floor(cell.X) + 1);
            int top = Math.Max(0, (int)Math.Floor(cell.Y) + 1);
            int right = Math.Min(image.Width, (int)Math.Ceiling(cell.X + cell.Width) - 1);
            int bottom = Math.Min(image.Height, (int)Math.Ceiling(cell.Y + cell.Height) - 1);
            int minY = int.MaxValue;
            int maxY = int.MinValue;
            for (int y = top; y < bottom; y++) {
                for (int x = left; x < right; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.A > 180 && pixel.R < 90 && pixel.G < 90 && pixel.B < 90) {
                        minY = Math.Min(minY, y);
                        maxY = Math.Max(maxY, y);
                    }
                }
            }

            Assert.NotEqual(int.MaxValue, minY);
            return (minY, maxY);
        }

        private static (int MinX, int MaxX) DarkPixelXExtent(OfficeRasterImage image, ExcelVisualCell cell) {
            int left = Math.Max(0, (int)Math.Floor(cell.X) + 1);
            int top = Math.Max(0, (int)Math.Floor(cell.Y) + 1);
            int right = Math.Min(image.Width, (int)Math.Ceiling(cell.X + cell.Width) - 1);
            int bottom = Math.Min(image.Height, (int)Math.Ceiling(cell.Y + cell.Height) - 1);
            int minX = int.MaxValue;
            int maxX = int.MinValue;
            for (int y = top; y < bottom; y++) {
                for (int x = left; x < right; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.A > 180 && pixel.R < 90 && pixel.G < 90 && pixel.B < 90) {
                        minX = Math.Min(minX, x);
                        maxX = Math.Max(maxX, x);
                    }
                }
            }

            Assert.NotEqual(int.MaxValue, minX);
            return (minX, maxX);
        }

        private static bool ContainsDarkPixel(OfficeRasterImage image, ExcelVisualCell cell) {
            int left = Math.Max(0, (int)Math.Floor(cell.X) + 1);
            int top = Math.Max(0, (int)Math.Floor(cell.Y) + 1);
            int right = Math.Min(image.Width, (int)Math.Ceiling(cell.X + cell.Width) - 1);
            int bottom = Math.Min(image.Height, (int)Math.Ceiling(cell.Y + cell.Height) - 1);
            for (int y = top; y < bottom; y++) {
                for (int x = left; x < right; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.A > 180 && pixel.R < 90 && pixel.G < 90 && pixel.B < 90) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ContainsVisiblePixel(OfficeRasterImage image, double left, double top, double right, double bottom) {
            int minX = Math.Max(0, (int)Math.Floor(left));
            int minY = Math.Max(0, (int)Math.Floor(top));
            int maxX = Math.Min(image.Width - 1, (int)Math.Ceiling(right));
            int maxY = Math.Min(image.Height - 1, (int)Math.Ceiling(bottom));
            if (minX > maxX || minY > maxY) {
                return false;
            }

            for (int y = minY; y <= maxY; y++) {
                for (int x = minX; x <= maxX; x++) {
                    OfficeColor pixel = image.GetPixel(x, y);
                    if (pixel.A > 180 && (pixel.R < 245 || pixel.G < 245 || pixel.B < 245)) {
                        return true;
                    }
                }
            }

            return false;
        }

        private static bool ContainsPixelNear(OfficeRasterImage image, double left, double top, double right, double bottom, OfficeColor expected, int tolerance) {
            int minX = Math.Max(0, (int)Math.Floor(left));
            int minY = Math.Max(0, (int)Math.Floor(top));
            int maxX = Math.Min(image.Width - 1, (int)Math.Ceiling(right));
            int maxY = Math.Min(image.Height - 1, (int)Math.Ceiling(bottom));
            if (minX > maxX || minY > maxY) {
                return false;
            }

            for (int y = minY; y <= maxY; y++) {
                for (int x = minX; x <= maxX; x++) {
                    OfficeColor actual = image.GetPixel(x, y);
                    if (actual.A > 180
                        && Math.Abs(actual.R - expected.R) <= tolerance
                        && Math.Abs(actual.G - expected.G) <= tolerance
                        && Math.Abs(actual.B - expected.B) <= tolerance) {
                        return true;
                    }
                }
            }

            return false;
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

        private static double ExtractFirstSvgFontSize(string svg) {
            const string marker = "font-size=\"";
            int start = svg.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(start >= 0, "SVG did not contain a font-size attribute.");
            start += marker.Length;
            int end = svg.IndexOf('"', start);
            Assert.True(end > start, "SVG font-size attribute was malformed.");
            string value = svg.Substring(start, end - start);
            return double.Parse(value, System.Globalization.CultureInfo.InvariantCulture);
        }

        private static double ExtractSvgClipWidth(string svg, string clipId) {
            return ExtractSvgClipDoubleAttribute(svg, clipId, "width");
        }

        private static double ExtractSvgClipX(string svg, string clipId) {
            return ExtractSvgClipDoubleAttribute(svg, clipId, "x");
        }

        private static double ExtractSvgTextX(string svg, string text) {
            string textMarker = ">" + text + "</text>";
            int textEnd = svg.IndexOf(textMarker, StringComparison.Ordinal);
            Assert.True(textEnd >= 0, "SVG did not contain text '" + text + "'.");
            int textStart = svg.LastIndexOf("<text", textEnd, StringComparison.Ordinal);
            Assert.True(textStart >= 0, "SVG text '" + text + "' did not have an opening text element.");
            return ExtractSvgElementDoubleAttribute(svg, textStart, "x");
        }

        private static double ExtractSvgClipDoubleAttribute(string svg, string clipId, string attributeName) {
            string marker = "id=\"" + clipId + "\"><rect";
            int clipStart = svg.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(clipStart >= 0, "SVG did not contain clip path '" + clipId + "'.");
            return ExtractSvgElementDoubleAttribute(svg, clipStart, attributeName);
        }

        private static double ExtractSvgElementDoubleAttribute(string svg, int elementStart, string attributeName) {
            string attributeMarker = attributeName + "=\"";
            int valueStart = svg.IndexOf(attributeMarker, elementStart, StringComparison.Ordinal);
            Assert.True(valueStart >= 0, "SVG element did not contain a " + attributeName + " attribute.");
            valueStart += attributeMarker.Length;
            int valueEnd = svg.IndexOf('"', valueStart);
            Assert.True(valueEnd > valueStart, "SVG element " + attributeName + " attribute was malformed.");
            return double.Parse(svg.Substring(valueStart, valueEnd - valueStart), System.Globalization.CultureInfo.InvariantCulture);
        }

        private sealed class SolidImageCodec : IOfficeRasterImageCodec {
            private readonly OfficeColor _color;

            internal SolidImageCodec(OfficeColor color) {
                _color = color;
            }

            internal int DecodeCalls { get; private set; }

            public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
                DecodeCalls++;
                image = new OfficeRasterImage(2, 2, _color);
                return true;
            }
        }

        private sealed class LengthOnlyImageStream : Stream {
            private long _position;

            internal LengthOnlyImageStream(long length) {
                Length = length;
            }

            internal int ReadCount { get; private set; }
            public override bool CanRead => true;
            public override bool CanSeek => true;
            public override bool CanWrite => false;
            public override long Length { get; }
            public override long Position { get => _position; set => _position = value; }
            public override void Flush() { }
            public override int Read(byte[] buffer, int offset, int count) {
                ReadCount++;
                return 0;
            }
            public override long Seek(long offset, SeekOrigin origin) {
                _position = origin == SeekOrigin.Begin
                    ? offset
                    : origin == SeekOrigin.Current
                        ? _position + offset
                        : Length + offset;
                return _position;
            }
            public override void SetLength(long value) => throw new NotSupportedException();
            public override void Write(byte[] buffer, int offset, int count) => throw new NotSupportedException();
        }

        private sealed class CountingVisualCellList : IReadOnlyList<ExcelVisualCell> {
            internal CountingVisualCellList(int count) {
                Count = count;
            }

            public int Count { get; }
            internal int ReadCount { get; private set; }

            public ExcelVisualCell this[int index] {
                get {
                    ReadCount++;
                    throw new InvalidOperationException("The oversized conditional-formatting guard must run before cell access.");
                }
            }

            public IEnumerator<ExcelVisualCell> GetEnumerator() {
                for (int index = 0; index < Count; index++) yield return this[index];
            }

            System.Collections.IEnumerator System.Collections.IEnumerable.GetEnumerator() => GetEnumerator();
        }
    }
}
