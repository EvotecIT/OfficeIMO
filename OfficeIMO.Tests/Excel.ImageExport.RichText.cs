using System.Text;
using System.Text.RegularExpressions;
using OfficeIMO.Drawing;
using OfficeIMO.Excel;
using Xunit;

namespace OfficeIMO.Tests {
    public class ExcelImageExportRichTextTests {
        [Fact]
        public void ExcelRange_ImageExportPreservesSingleLineRichTextRunsInSvg() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Rich");
            sheet.SetColumnWidth(1, 26);
            sheet.SetRowHeight(1, 28);
            sheet.CellAt(1, 1).SetRichText(
                new ExcelRichTextRun("Strong") { Bold = true, FontColor = "FF0000" },
                new ExcelRichTextRun(" note") { Italic = true, Underline = true, FontColor = "0563C1", FontSize = 13D });

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = Encoding.UTF8.GetString(svgResult.Bytes);

            Assert.Equal(2, snapshot.Cells[0].RichTextRuns.Count);
            Assert.Equal("Strong", snapshot.Cells[0].RichTextRuns[0].Text);
            Assert.Equal(" note", snapshot.Cells[0].RichTextRuns[1].Text);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.Contains("#FF0000", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#0563C1", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("text-anchor=\"start\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline\"", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesHardBreakRichTextRunsInPngAndSvg() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Rich");
            sheet.SetColumnWidth(1, 18);
            sheet.SetRowHeight(1, 44);
            sheet.CellAt(1, 1)
                .SetRichText(
                    new ExcelRichTextRun("Wrapped") { Bold = true, FontColor = "FF0000" },
                    new ExcelRichTextRun("\ntext") { Italic = true, FontColor = "0563C1" });

            ExcelRange range = sheet.Range("A1:A1");
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = Encoding.UTF8.GetString(svgResult.Bytes);

            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.Contains("Wrapped", svg, StringComparison.Ordinal);
            Assert.Contains("text", svg, StringComparison.Ordinal);
            Assert.Contains("#FF0000", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#0563C1", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesShrinkToFitRichTextRunsInPngAndSvg() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Rich");
            sheet.SetColumnWidth(1, 8);
            sheet.SetRowHeight(1, 30);
            sheet.CellAt(1, 1)
                .SetShrinkToFit()
                .SetRichText(
                    new ExcelRichTextRun("Shrink") { Bold = true, FontColor = "FF0000" },
                    new ExcelRichTextRun(" text") { Italic = true, FontColor = "0563C1", FontSize = 18D });

            ExcelRange range = sheet.Range("A1:A1");
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = Encoding.UTF8.GetString(svgResult.Bytes);

            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.Contains("Shrink", svg, StringComparison.Ordinal);
            Assert.Contains("text", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            MatchCollection fontSizeMatches = Regex.Matches(svg, "font-size=\"([0-9.]+)\"");
            Assert.Contains(fontSizeMatches.Cast<Match>(), match => double.Parse(match.Groups[1].Value, System.Globalization.CultureInfo.InvariantCulture) < 18D);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesRotatedRichTextRunsWithApproximationDiagnostic() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorkSheet("Rich");
            sheet.SetColumnWidth(1, 16);
            sheet.SetRowHeight(1, 60);
            sheet.CellAt(1, 1)
                .SetTextRotation(45)
                .SetRichText(
                    new ExcelRichTextRun("Tilt") { Bold = true, FontColor = "0F766E", FontSize = 14D },
                    new ExcelRichTextRun(" rich") { Italic = true, FontColor = "7C3AED", FontSize = 13D },
                    new ExcelRichTextRun(" text") { Underline = true, FontColor = "2563EB", FontSize = 13D });

            ExcelRange range = sheet.Range("A1:A1");
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = Encoding.UTF8.GetString(svgResult.Bytes);

            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.Contains(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextRotationApproximation && diagnostic.Source == "Rich!A1");
            Assert.Contains(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextRotationApproximation && diagnostic.Source == "Rich!A1");
            Assert.Contains("transform=\"rotate(-45", svg, StringComparison.Ordinal);
            Assert.Contains("Tilt", svg, StringComparison.Ordinal);
            Assert.Contains("rich", svg, StringComparison.Ordinal);
            Assert.Contains("text", svg, StringComparison.Ordinal);
            Assert.Contains("#0F766E", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#7C3AED", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2563EB", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline\"", svg, StringComparison.Ordinal);
        }
    }
}
