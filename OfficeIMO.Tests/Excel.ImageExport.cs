using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
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
            ExcelSheet sheet = document.AddWorkSheet("Data");
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
        public void ExcelRange_ImageExportUsesNumberFormatLiteralsAndSections() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Formats");
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
        public void ExcelRange_ImageExportLaysOutMultilineCellTextThroughSharedDrawingLayout() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Wrap");
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
        public void ExcelRange_ImageExportResolvesThemeTintColorsAcrossSnapshotsAndRenderers() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Theme");
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
            Assert.Equal("FF602827", visualStyle.FontColorArgb);
            Assert.Equal("FF9BBB59", visualStyle.Border!.Top!.ColorArgb);
            Assert.Equal(visualStyle.FillColorArgb, directStyle.FillColorArgb);
            Assert.Equal(visualStyle.FontColorArgb, directStyle.FontColorArgb);
            Assert.Equal(visualStyle.FillColorArgb, inspectedCell.Style!.FillColorArgb);
            Assert.Contains("#95B3D7", svg, StringComparison.Ordinal);
            Assert.Contains("#602827", svg, StringComparison.Ordinal);
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
        public void ExcelRange_ImageExportAppliesConditionalColorScalesAndDataBarsWithDiagnostics() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Conditional");
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
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ConditionalIconSetUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Conditional!C1:C3", diagnostic.Source);
            Assert.Contains("#FF0000", svg, StringComparison.Ordinal);
            Assert.Contains("#0000FF", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            ExcelVisualCell firstScaleCell = snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1);
            AssertPixelNear(
                rendered!,
                (int)(firstScaleCell.X + firstScaleCell.Width - 8),
                (int)(firstScaleCell.Y + firstScaleCell.Height - 8),
                OfficeColor.Red,
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(finalBar.X + finalBar.Width - 8),
                (int)(finalBar.Y + (finalBar.Height / 2D)),
                OfficeColor.Blue,
                tolerance: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesConditionalCellIsAndFormulaDifferentialFills() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Rules");
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
            Assert.Equal("FFFF0000", snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 2).Style.FillColorArgb);
            Assert.Equal("FF00FF00", snapshot.Cells.Single(cell => cell.Row == 3 && cell.Column == 2).Style.FillColorArgb);
            Assert.Contains("#C6EFCE", svg, StringComparison.Ordinal);
            Assert.Contains("#FEE2E2", svg, StringComparison.Ordinal);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            ExcelVisualCell cellIsRendered = snapshot.Cells.Single(cell => cell.Row == 2 && cell.Column == 1);
            ExcelVisualCell formulaRendered = snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 2);
            AssertPixelNear(
                rendered!,
                (int)(cellIsRendered.X + cellIsRendered.Width - 8),
                (int)(cellIsRendered.Y + cellIsRendered.Height - 8),
                OfficeColor.FromRgb(198, 239, 206),
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(formulaRendered.X + formulaRendered.Width - 8),
                (int)(formulaRendered.Y + formulaRendered.Height - 8),
                OfficeColor.FromRgb(254, 226, 226),
                tolerance: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnsupportedConditionalRuleShapes() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Unsupported");
            sheet.CellValue(1, 1, "Hot");
            sheet.CellValue(2, 1, "Warm");
            sheet.CellValue(1, 2, 2);
            sheet.CellValue(2, 2, 4);
            sheet.SetColumnWidth(1, 12);
            sheet.SetColumnWidth(2, 12);
            sheet.AddConditionalRule("A1:A2", ConditionalFormattingOperatorValues.Equal, "\"Hot\"", fillColor: "FEE2E2");
            sheet.AddConditionalFormulaRule("B1:B2", "MOD(B1,2)=0", fillColor: "C6EFCE");
            AddUnsupportedTimePeriodRule(sheet, "B1:B2");

            ExcelRange range = sheet.Range("A1:B2");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 1).Style.FillColorArgb);
            Assert.Null(snapshot.Cells.Single(cell => cell.Row == 1 && cell.Column == 2).Style.FillColorArgb);
            OfficeImageExportDiagnostic cellIs = Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalCellIsUnsupported);
            OfficeImageExportDiagnostic formula = Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalFormulaUnsupported);
            OfficeImageExportDiagnostic rule = Assert.Single(png.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ConditionalRuleUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, cellIs.Severity);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, formula.Severity);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, rule.Severity);
            Assert.Equal("Unsupported!A1:A2", cellIs.Source);
            Assert.Equal("Unsupported!B1:B2", formula.Source);
            Assert.Equal("Unsupported!B1:B2", rule.Source);
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesTextConditionalFills() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("TextRules");
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
            ExcelSheet sheet = document.AddWorkSheet("Averages");
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
            AssertPixelNear(
                rendered!,
                (int)(aboveCell.X + aboveCell.Width - 8),
                (int)(aboveCell.Y + aboveCell.Height - 8),
                OfficeColor.FromRgb(198, 239, 206),
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(belowCell.X + belowCell.Width - 8),
                (int)(belowCell.Y + belowCell.Height - 8),
                OfficeColor.FromRgb(219, 234, 254),
                tolerance: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportAppliesDuplicateAndUniqueValueConditionalFills() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Distinct");
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
        public void ExcelRange_ImageExportAppliesTopBottomConditionalFillsIncludingPercentVariants() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("TopBottom");
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
            AssertPixelNear(
                rendered!,
                (int)(topCell.X + topCell.Width - 8),
                (int)(topCell.Y + topCell.Height - 8),
                OfficeColor.FromRgb(198, 239, 206),
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(bottomCell.X + bottomCell.Width - 8),
                (int)(bottomCell.Y + bottomCell.Height - 8),
                OfficeColor.FromRgb(254, 226, 226),
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(topPercentCell.X + topPercentCell.Width - 8),
                (int)(topPercentCell.Y + topPercentCell.Height - 8),
                OfficeColor.FromRgb(252, 228, 214),
                tolerance: 3);
            AssertPixelNear(
                rendered!,
                (int)(bottomPercentCell.X + bottomPercentCell.Width - 8),
                (int)(bottomPercentCell.Y + bottomPercentCell.Height - 8),
                OfficeColor.FromRgb(219, 234, 254),
                tolerance: 3);
        }

        [Fact]
        public void ExcelSheet_ExportsUsedRangeToPngAndSvg() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Summary");
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
            ExcelSheet sheet = document.AddWorkSheet("Images");
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
        public void ExcelRange_ImageExportIncludesAndClipsImagesOverlappingSelectedRange() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("ImageClip");
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
            Assert.True(image.X < 0D, "The overlapping image should keep its true negative X position relative to the exported range.");
            Assert.DoesNotContain(png.Diagnostics, diagnostic => diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error);
            Assert.Contains("clip-path=\"url(#xl-image-clip-", svg, StringComparison.Ordinal);
            Assert.Contains("x=\"-", svg, StringComparison.Ordinal);
            Assert.True(OfficePngReader.TryDecode(png.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            AssertPixelNear(rendered!, 8, 8, OfficeColor.FromRgb(220, 38, 38), tolerance: 3);
        }

        [Fact]
        public void ExcelRange_ImageExportUsesTwoCellImageAnchorDimensions() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            byte[] banner = CreateSolidPng(32, 32, OfficeColor.FromRgb(37, 99, 235));
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("TwoCell");
                sheet.SetColumnWidth(1, 10);
                sheet.SetColumnWidth(2, 10);
                sheet.SetColumnWidth(3, 10);
                sheet.SetColumnWidth(4, 10);
                sheet.SetRowHeight(1, 24);
                sheet.SetRowHeight(2, 24);
                sheet.SetRowHeight(3, 24);
                sheet.CellValue(1, 1, "Two-cell");
                document.Save(false);
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
        public void ExcelRange_ImageExportHonorsPictureCropRectangle() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            byte[] croppedSource = CreateHorizontalCropPng();
            using (ExcelDocument document = ExcelDocument.Create(filePath)) {
                ExcelSheet sheet = document.AddWorkSheet("Crop");
                sheet.SetColumnWidth(1, 14);
                sheet.SetColumnWidth(2, 14);
                sheet.SetRowHeight(1, 30);
                sheet.SetRowHeight(2, 30);
                document.Save(false);
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
                ExcelSheet sheet = document.AddWorkSheet("RotateImage");
                for (int column = 1; column <= 4; column++) {
                    sheet.SetColumnWidth(column, 14);
                }

                for (int row = 1; row <= 5; row++) {
                    sheet.SetRowHeight(row, 30);
                }

                document.Save(false);
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
                ExcelSheet sheet = document.AddWorkSheet("TransformImage");
                for (int column = 1; column <= 5; column++) {
                    sheet.SetColumnWidth(column, 14);
                }

                for (int row = 1; row <= 6; row++) {
                    sheet.SetRowHeight(row, 30);
                }

                document.Save(false);
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
        public void ExcelRange_ImageExportEmbedsJpegInSvgAndReportsPngRasterLimitation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Jpeg");
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
            Assert.DoesNotContain(svg.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.ImageSvgFormatUnsupported);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.ImageRasterFormatUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Jpeg!PhotoJpeg", diagnostic.Source);
        }

        [Fact]
        public void ExcelRange_ExportsChartsThroughSharedDrawingRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Charts");
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
            ExcelSheet sheet = document.AddWorkSheet("ChartStyle");
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
            ExcelSheet sheet = document.AddWorkSheet("ChartAreaLines");
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
            ExcelSheet sheet = document.AddWorkSheet("ChartTitle");
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
            ExcelSheet sheet = document.AddWorkSheet("ChartTitleTypography");
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
            ExcelSheet sheet = document.AddWorkSheet("ChartTitleStyle");
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
            ExcelSheet sheet = document.AddWorkSheet("ChartSeries");
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
            ExcelSheet sheet = document.AddWorkSheet("ChartPoints");
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
            ExcelSheet sheet = document.AddWorkSheet("ChartMarkers");
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
        public void ExcelRange_ImageExportCarriesChartMarkerSizeIntoSharedRenderer() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("MarkerSize");
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
            ExcelSheet sheet = document.AddWorkSheet("MarkerShape");
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
            ExcelSheet sheet = document.AddWorkSheet("MarkerX");
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
            ExcelSheet sheet = document.AddWorkSheet("MarkerSymbols");
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
            ExcelSheet sheet = document.AddWorkSheet("MarkerOutline");
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
            ExcelSheet sheet = document.AddWorkSheet("SeriesLineWidth");
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
            ExcelSheet sheet = document.AddWorkSheet("SeriesLineDash");
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
            ExcelSheet sheet = document.AddWorkSheet("ChartGridlines");
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
            ExcelSheet sheet = document.AddWorkSheet("MinorGridlines");
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
            ExcelSheet sheet = document.AddWorkSheet("GridlineWidth");
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
            ExcelSheet sheet = document.AddWorkSheet("GridlineDash");
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
            ExcelSheet sheet = document.AddWorkSheet("NoGridlines");
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
            ExcelSheet sheet = document.AddWorkSheet("AxisColor");
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
            ExcelSheet sheet = document.AddWorkSheet("NoAxisLines");
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
            ExcelSheet sheet = document.AddWorkSheet("AxisWidth");
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
            ExcelSheet sheet = document.AddWorkSheet("AxisDash");
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
            ExcelSheet sheet = document.AddWorkSheet("Trendline");
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
            ExcelSheet first = document.AddWorkSheet("First");
            first.CellValue(1, 1, "One");
            ExcelSheet second = document.AddWorkSheet("Second");
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
        public void ExcelRange_ImageExportHonorsCellVerticalTextAlignment() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Align");
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
            ExcelSheet sheet = document.AddWorkSheet("DefaultAlign");
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
            ExcelSheet sheet = document.AddWorkSheet("Wrap");
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
            ExcelSheet sheet = document.AddWorkSheet("Clip");
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
            ExcelSheet sheet = document.AddWorkSheet("FontSize");
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
            ExcelSheet sheet = document.AddWorkSheet("Shrink");
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
            ExcelSheet sheet = document.AddWorkSheet("Rotate");
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
            ExcelSheet sheet = document.AddWorkSheet("ClipRotate");
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
        public void ExcelRange_ImageExportReportsUnsupportedStackedTextRotation() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Stacked");
            sheet.CellValue(1, 1, "Stacked");
            sheet.CellAt(1, 1).SetTextRotation(255);

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult png = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            Assert.Equal(255, snapshot.Cells[0].Style.TextRotation);
            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellStackedTextRotationUnsupported);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("Stacked!A1", diagnostic.Source);
        }

        private static byte[] CreateSolidPng(int width, int height, OfficeColor color) {
            OfficeRasterImage image = new OfficeRasterImage(width, height, OfficeColor.Transparent);
            image.Fill(color);
            return OfficePngWriter.Encode(image);
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

        private static void AddUnsupportedTimePeriodRule(ExcelSheet sheet, string range) {
            Worksheet worksheet = sheet.WorksheetPart.Worksheet ?? throw new InvalidOperationException("Worksheet is missing.");
            var conditional = new ConditionalFormatting {
                SequenceOfReferences = new ListValue<StringValue> { InnerText = range }
            };
            conditional.Append(new ConditionalFormattingRule {
                Type = ConditionalFormatValues.TimePeriod,
                Priority = 99
            });

            SheetData? sheetData = worksheet.GetFirstChild<SheetData>();
            if (sheetData == null) {
                worksheet.Append(conditional);
            } else {
                worksheet.InsertAfter(conditional, sheetData);
            }

            worksheet.Save();
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

        private static void AddTwoCellAnchoredImage(string filePath, byte[] imageBytes) {
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
                    new Xdr.ColumnOffset("0"),
                    new Xdr.RowId("0"),
                    new Xdr.RowOffset("0")),
                new Xdr.ToMarker(
                    new Xdr.ColumnId("3"),
                    new Xdr.ColumnOffset("0"),
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

        private static void AssertPixelNear(OfficeRasterImage image, int x, int y, OfficeColor expected, int tolerance) {
            OfficeColor actual = image.GetPixel(x, y);
            Assert.True(
                Math.Abs(actual.R - expected.R) <= tolerance
                && Math.Abs(actual.G - expected.G) <= tolerance
                && Math.Abs(actual.B - expected.B) <= tolerance,
                $"Expected pixel {x},{y} near {expected}, got {actual}.");
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
    }
}
