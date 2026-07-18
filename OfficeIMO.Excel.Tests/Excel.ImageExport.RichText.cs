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
            ExcelSheet sheet = document.AddWorksheet("Rich");
            sheet.SetColumnWidth(1, 26);
            sheet.SetRowHeight(1, 28);
            sheet.CellAt(1, 1).SetRichText(
                new ExcelRichTextRun("Strong") { Bold = true, FontColor = "FF0000" },
                new ExcelRichTextRun(" note") { Italic = true, Underline = true, FontColor = "0563C1", FontSize = 13D },
                new ExcelRichTextRun(" gone") { Strikethrough = true, FontColor = "6B7280", FontSize = 12D });

            ExcelRange range = sheet.Range("A1:A1");
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot();
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = Encoding.UTF8.GetString(svgResult.Bytes);

            Assert.Equal(3, snapshot.Cells[0].RichTextRuns.Count);
            Assert.Equal("Strong", snapshot.Cells[0].RichTextRuns[0].Text);
            Assert.Equal(" note", snapshot.Cells[0].RichTextRuns[1].Text);
            Assert.True(snapshot.Cells[0].RichTextRuns[2].Strikethrough);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.Contains("#FF0000", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#0563C1", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#6B7280", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("text-anchor=\"start\"", svg, StringComparison.Ordinal);
            Assert.Contains("xml:space=\"preserve\"", svg, StringComparison.Ordinal);
            Assert.Contains("> note</text>", svg, StringComparison.Ordinal);
            Assert.Contains("> gone</text>", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"line-through\"", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportReportsUnresolvedRichTextRunFontFamilyFallback() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("RichFont");
            sheet.SetColumnWidth(1, 22);
            sheet.SetRowHeight(1, 28);
            sheet.CellAt(1, 1).SetRichText(
                new ExcelRichTextRun("Missing run font") {
                    FontName = "OfficeIMO Missing Rich Font",
                    FontColor = "0F766E"
                });

            OfficeImageExportResult png = sheet.Range("A1:A1").ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });

            OfficeImageExportDiagnostic diagnostic = Assert.Single(png.Diagnostics, item => item.Code == OfficeImageExportDiagnosticCodes.FontSubstituted);
            Assert.Equal(OfficeImageExportDiagnosticSeverity.Warning, diagnostic.Severity);
            Assert.Equal("RichFont!A1", diagnostic.Source);
            Assert.Contains("OfficeIMO Missing Rich Font", diagnostic.Message);
            Assert.DoesNotContain(png.Diagnostics, item => item.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesHardBreakRichTextRunsInPngAndSvg() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Rich");
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
            ExcelSheet sheet = document.AddWorksheet("Rich");
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
        public void ExcelRange_ImageExportClipsOverflowingRichTextWithoutInventingEllipsis() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("RichClip");
            sheet.SetColumnWidth(1, 8);
            sheet.SetRowHeight(1, 24);
            sheet.CellAt(1, 1).SetRichText(
                new ExcelRichTextRun("Overflowing") { Bold = true, FontColor = "DC2626", FontSize = 12D },
                new ExcelRichTextRun(" rich text should clip") { Italic = true, FontColor = "2563EB", FontSize = 12D });

            ExcelRange range = sheet.Range("A1:A1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, options);
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            string svg = Encoding.UTF8.GetString(svgResult.Bytes);

            Assert.Contains(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "RichClip!A1");
            Assert.Contains(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "RichClip!A1");
            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.Contains("Overflowing", svg, StringComparison.Ordinal);
            Assert.Contains(" rich text should clip", svg, StringComparison.Ordinal);
            Assert.DoesNotContain("...", svg, StringComparison.Ordinal);
            Assert.Contains("clip-path=\"url(#xl-text-1-1)\"", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportSpillsRichTextIntoBlankNeighborCells() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("RichSpill");
            sheet.SetColumnWidth(1, 6);
            sheet.SetColumnWidth(2, 24);
            sheet.SetColumnWidth(3, 8);
            sheet.SetRowHeight(1, 26);
            sheet.CellAt(1, 1).SetRichText(
                new ExcelRichTextRun("Rich") { Bold = true, FontColor = "0F766E", FontSize = 12D },
                new ExcelRichTextRun(" text spills") { Italic = true, FontColor = "7C3AED", FontSize = 12D });
            sheet.CellValue(1, 3, "Stop");

            ExcelRange range = sheet.Range("A1:C1");
            ExcelImageExportOptions options = new() { ShowGridlines = false };
            ExcelRangeVisualSnapshot snapshot = range.CreateVisualSnapshot(options);
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, options);
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, options);
            string svg = Encoding.UTF8.GetString(svgResult.Bytes);

            ExcelVisualCell first = snapshot.Cells.Single(cell => cell.Column == 1);
            ExcelVisualCell blankNeighbor = snapshot.Cells.Single(cell => cell.Column == 2);
            double expectedWidth = first.Width + blankNeighbor.Width;
            Assert.Equal(expectedWidth, ExtractSvgClipWidth(svg, "xl-text-1-1"), precision: 2);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "RichSpill!A1");
            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextClipped && diagnostic.Source == "RichSpill!A1");
            Assert.Contains("Rich", svg, StringComparison.Ordinal);
            Assert.Contains("text spills", svg, StringComparison.Ordinal);
            Assert.Contains("Stop", svg, StringComparison.Ordinal);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
        }

        [Fact]
        public void ExcelRange_ImageExportPreservesRotatedRichTextRunsWithApproximationDiagnostic() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Rich");
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

        [Fact]
        public void ExcelRange_ImageExportPreservesStackedRichTextRunsWithApproximationDiagnostic() {
            string filePath = Path.Combine(Path.GetTempPath(), Guid.NewGuid() + ".xlsx");
            using ExcelDocument document = ExcelDocument.Create(filePath);
            ExcelSheet sheet = document.AddWorksheet("Rich");
            sheet.SetColumnWidth(1, 5);
            sheet.SetRowHeight(1, 96);
            sheet.CellAt(1, 1)
                .SetTextRotation(255)
                .SetRichText(
                    new ExcelRichTextRun("S") { Bold = true, FontColor = "0F766E", FontSize = 12D },
                    new ExcelRichTextRun("V") { Italic = true, FontColor = "7C3AED", FontSize = 12D },
                    new ExcelRichTextRun("G") { Underline = true, FontColor = "2563EB", FontSize = 12D });

            ExcelRange range = sheet.Range("A1:A1");
            OfficeImageExportResult pngResult = range.ExportImage(OfficeImageExportFormat.Png, new ExcelImageExportOptions { ShowGridlines = false });
            OfficeImageExportResult svgResult = range.ExportImage(OfficeImageExportFormat.Svg, new ExcelImageExportOptions { ShowGridlines = false });
            string svg = Encoding.UTF8.GetString(svgResult.Bytes);

            Assert.True(OfficePngReader.TryDecode(pngResult.Bytes, out OfficeRasterImage? rendered));
            Assert.NotNull(rendered);
            Assert.DoesNotContain(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.DoesNotContain(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellRichTextLayoutApproximation);
            Assert.Contains(pngResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextRotationApproximation && diagnostic.Source == "Rich!A1");
            Assert.Contains(svgResult.Diagnostics, diagnostic => diagnostic.Code == ExcelImageExportDiagnosticCodes.CellTextRotationApproximation && diagnostic.Source == "Rich!A1");
            Assert.DoesNotContain("rotate(", svg, StringComparison.Ordinal);
            Assert.Contains(">S</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">V</text>", svg, StringComparison.Ordinal);
            Assert.Contains(">G</text>", svg, StringComparison.Ordinal);
            Assert.Contains("#0F766E", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#7C3AED", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("#2563EB", svg, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("font-weight=\"700\"", svg, StringComparison.Ordinal);
            Assert.Contains("font-style=\"italic\"", svg, StringComparison.Ordinal);
            Assert.Contains("text-decoration=\"underline\"", svg, StringComparison.Ordinal);
        }

        private static double ExtractSvgClipWidth(string svg, string clipId) {
            string marker = "id=\"" + clipId + "\"><rect";
            int clipStart = svg.IndexOf(marker, StringComparison.Ordinal);
            Assert.True(clipStart >= 0, "SVG did not contain clip path '" + clipId + "'.");
            int widthStart = svg.IndexOf("width=\"", clipStart, StringComparison.Ordinal);
            Assert.True(widthStart >= 0, "SVG clip path '" + clipId + "' did not contain a width attribute.");
            widthStart += "width=\"".Length;
            int widthEnd = svg.IndexOf('"', widthStart);
            Assert.True(widthEnd > widthStart, "SVG clip path '" + clipId + "' width attribute was malformed.");
            return double.Parse(svg.Substring(widthStart, widthEnd - widthStart), System.Globalization.CultureInfo.InvariantCulture);
        }
    }
}
