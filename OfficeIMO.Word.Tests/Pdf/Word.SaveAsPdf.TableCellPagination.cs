using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("docx", WordCompatibilityMode.Word2013, false, false, true)]
    [InlineData("docx", WordCompatibilityMode.Word2013, true, false, false)]
    [InlineData("docx", WordCompatibilityMode.Word2013, false, true, false)]
    [InlineData("docx", WordCompatibilityMode.Word2010, true, false, true)]
    [InlineData("docx", WordCompatibilityMode.Word2007, false, true, true)]
    [InlineData("doc", WordCompatibilityMode.Word2013, false, false, true)]
    [InlineData("doc", WordCompatibilityMode.Word2013, true, false, true)]
    [InlineData("doc", WordCompatibilityMode.Word2013, false, true, true)]
    public void SaveAsPdf_CellParagraphPaginationUsesTheSourceLayoutMode(string extension, WordCompatibilityMode mode,
        bool widowControl, bool keepTogether, bool firstLineFits) {
        string source = Path.Combine(_directoryWithFiles, $"CellPagination-{extension}-{mode}-{widowControl}-{keepTogether}." + extension);
        string target = source + ".pdf";
        using (WordDocument document = WordDocument.Create(source)) {
            document.CompatibilitySettings.CompatibilityMode = mode;
            WordTable table = CreatePaginationTable(document, 8);
            foreach (WordTableRow row in table.Rows) {
                WordParagraph paragraph = row.Cells[0].Paragraphs[0];
                paragraph.AvoidWidowAndOrphanOverride = widowControl;
                paragraph.KeepLinesTogetherOverride = keepTogether;
            }
            document.Save();
        }
        using (WordDocument document = WordDocument.Load(source)) {
            WordParagraph paragraph = document.Tables[0].Rows[5].Cells[0].Paragraphs[0];
            Assert.Equal(widowControl, paragraph.AvoidWidowAndOrphanOverride);
            Assert.Equal(keepTogether, paragraph.KeepLinesTogetherOverride);
            document.SaveAsPdf(target, CellPaginationOptions());
        }
        using var pdf = PdfPigDocument.Open(target);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(FindWordStartX(pdf.GetPage(1), "Row06Line001"), firstLineFits ? 39.9 : 259.9, firstLineFits ? 40.1 : 260.1);
        Assert.Contains("Row08Line003", pdf.GetPage(1).Text);
    }

    [Theory]
    [InlineData("table", true)]
    [InlineData("paragraph", false)]
    [InlineData("direct", true)]
    [InlineData("conditional", true)]
    [InlineData("document", false)]
    public void SaveAsPdf_CellWidowControlResolvesTheFormattingHierarchy(string sourceLevel, bool firstLineFits) {
        string source = Path.Combine(_directoryWithFiles, "CellPaginationHierarchy-" + sourceLevel + ".docx");
        string target = source + ".pdf";
        using (WordDocument document = WordDocument.Create(source)) {
            WordTable table = CreatePaginationTable(document, 6);
            Styles styles = document._wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!;
            if (sourceLevel == "document") {
                var defaults = styles.DocDefaults!.GetFirstChild<ParagraphPropertiesDefault>()!;
                var properties = defaults.GetFirstChild<ParagraphPropertiesBaseStyle>()!;
                properties.RemoveAllChildren<WidowControl>();
                properties.Append(new WidowControl());
            } else {
                var tableStyle = new Style(new StyleName { Val = "Cell pagination" },
                    new StyleParagraphProperties(new WidowControl { Val = sourceLevel == "conditional" })) {
                    Type = StyleValues.Table, StyleId = "CellPaginationTable"
                };
                if (sourceLevel == "conditional") {
                    tableStyle.Append(new TableStyleProperties(new StyleParagraphProperties(new WidowControl { Val = false })) {
                        Type = TableStyleOverrideValues.LastRow
                    });
                    table.ConditionalFormattingLastRow = true;
                }
                styles.Append(tableStyle);
                table._tableProperties!.TableStyle = new TableStyle { Val = tableStyle.StyleId };
                if (sourceLevel is "paragraph" or "direct") {
                    styles.Append(new Style(new StyleName { Val = "Cell paragraph" }, new StyleParagraphProperties(new WidowControl())) {
                        Type = StyleValues.Paragraph, StyleId = "CellPaginationParagraph"
                    });
                    foreach (WordTableRow row in table.Rows) {
                        row.Cells[0].Paragraphs[0]._paragraph.ParagraphProperties!.ParagraphStyleId = new ParagraphStyleId { Val = "CellPaginationParagraph" };
                    }
                }
            }
            foreach (WordTableRow row in table.Rows) {
                row.Cells[0].Paragraphs[0].AvoidWidowAndOrphanOverride = sourceLevel == "direct" ? false : null;
            }
            document.Save();
            document.SaveAsPdf(target, CellPaginationOptions());
        }
        using var pdf = PdfPigDocument.Open(target);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(FindWordStartX(pdf.GetPage(1), "Row06Line001"), firstLineFits ? 39.9 : 259.9, firstLineFits ? 40.1 : 260.1);
    }

    [Theory]
    [InlineData(0, true)]
    [InlineData(2, false)]
    public void SaveAsPdf_ExplicitPdfTableRowGroupingRemainsConfigurable(int minimumRows, bool firstFrame) {
        string source = Path.Combine(_directoryWithFiles, "ConfiguredCellPagination-" + minimumRows + ".docx");
        string target = source + ".pdf";
        using (WordDocument document = WordDocument.Create(source)) {
            CreatePaginationTable(document, 6).Style = WordTableStyle.TableNormal;
            WordPdfExportOptions options = CellPaginationOptions();
            options.PdfOptions = new PdfCore.PdfCreateOptions { DefaultTableStyle = new PdfCore.PdfTableStyle {
                MinimumBodyRowsOnFirstPage = minimumRows, MinimumBodyRowsOnLastPage = minimumRows,
                HeaderRowCount = 0, CellPaddingX = global::OfficeIMO.OfficeLength.Points(0), CellPaddingY = global::OfficeIMO.OfficeLength.Points(0), SpacingBefore = global::OfficeIMO.OfficeLength.Points(0), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0)
            } };
            document.SaveAsPdf(target, options);
        }
        using var pdf = PdfPigDocument.Open(target);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(FindWordStartX(pdf.GetPage(1), "Row05Line001"), firstFrame ? 39.9 : 259.9, firstFrame ? 40.1 : 260.1);
    }

    // Word fills a breaking row cell by cell: a widow rule or a trailing paragraph
    // space that limits one cell does not cut the other cell's lines.
    [Theory]
    [InlineData(true, 10D, "A1 A2", "B1 B2 B3 B4")]
    [InlineData(false, 10D, "A1 A2 A3", "B1 B2 B3 B4")]
    public void SaveAsPdf_BreakingRowKeepsEachCellsFittingLinesOnThePage(bool widowControl, double spacingAfter,
        string firstPageA, string firstPageB) {
        using WordDocument document = CreateBreakingRowDocument(fillerLines: 12, widowControl, spacingAfter);

        using var pdf = PdfPigDocument.Open(document.ToPdf(CellPaginationOptions()).ToBytes());

        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal(firstPageA, string.Join(" ", BreakingRowLines(pdf.GetPage(1), "A")));
        Assert.Equal(firstPageB, string.Join(" ", BreakingRowLines(pdf.GetPage(1), "B")));
        Assert.Equal(new[] { "A1", "A2", "A3", "A4" }, BreakingRowLines(pdf.GetPage(1), "A").Concat(BreakingRowLines(pdf.GetPage(2), "A")));
        Assert.Equal(new[] { "B1", "B2", "B3", "B4", "B5", "B6" }, BreakingRowLines(pdf.GetPage(1), "B").Concat(BreakingRowLines(pdf.GetPage(2), "B")));
    }

    // Word applies a cell's vertical alignment where its row starts; the continued part starts at the top.
    [Fact]
    public void SaveAsPdf_BreakingRowContinuesCenteredCellsAtTheTop() {
        using WordDocument document = CreateBreakingRowDocument(fillerLines: 14, widowControl: true, spacingAfter: 0D);
        document.Tables[0].Rows[0].Cells[0].VerticalAlignment = WordTableVerticalAlignment.Center;

        using var pdf = PdfPigDocument.Open(document.ToPdf(CellPaginationOptions()).ToBytes());

        var page = pdf.GetPage(2);
        Assert.Equal(FindWordBaseline(page, "B3"), FindWordBaseline(page, "A3"), 2);
    }

    // Word repeats and keeps whole only header rows marked to repeat. A first row that a table
    // style merely formats as a header breaks across pages like any other row.
    [Fact]
    public void SaveAsPdf_StyledFirstRowThatDoesNotRepeatBreaksAcrossPages() {
        using WordDocument document = CreateBreakingRowDocument(fillerLines: 12, widowControl: false, spacingAfter: 0D);
        document.Tables[0].ConditionalFormattingFirstRow = true;

        using var pdf = PdfPigDocument.Open(document.ToPdf(CellPaginationOptions()).ToBytes());

        Assert.Equal("A1 A2 A3 A4", string.Join(" ", BreakingRowLines(pdf.GetPage(1), "A")));
        Assert.Equal("B5 B6", string.Join(" ", BreakingRowLines(pdf.GetPage(2), "B")));
    }

    private static WordDocument CreateBreakingRowDocument(int fillerLines, bool widowControl, double spacingAfter) {
        WordDocument document = WordDocument.Create();
        void Format(WordParagraph paragraph, double after) {
            paragraph.FontFamily = "Arial"; paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Points(20));
            paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero; paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(after);
            paragraph.AvoidWidowAndOrphanOverride = widowControl;
        }
        for (int line = 1; line <= fillerLines; line++) Format(document.AddParagraph("Filler" + line), 0D);
        WordTable table = document.AddTable(1, 2, WordTableStyle.TableGrid);
        table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Nil, global::OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeColor.Black);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.ConditionalFormattingFirstRow = false;
        for (int column = 0; column < 2; column++) {
            WordTableCell cell = table.Rows[0].Cells[column];
            cell.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Points(160));
            cell.MarginTop = global::OfficeIMO.OfficeLength.Zero; cell.MarginBottom = global::OfficeIMO.OfficeLength.Zero;
            string name = column == 0 ? "A" : "B";
            cell.Paragraphs[0].Text = string.Join("\n", Enumerable.Range(1, column == 0 ? 4 : 6).Select(line => name + line));
            Format(cell.Paragraphs[0], spacingAfter);
        }
        return document;
    }

    private static string[] BreakingRowLines(UglyToad.PdfPig.Content.Page page, string cell) =>
        page.GetWords().Where(word => word.Text.Length == 2 && word.Text.StartsWith(cell, StringComparison.Ordinal))
            .OrderByDescending(word => word.BoundingBox.Bottom).Select(word => word.Text).ToArray();

    private static double FindWordBaseline(UglyToad.PdfPig.Content.Page page, string text) =>
        page.GetWords().Single(word => word.Text == text).Letters[0].StartBaseLine.Y;

    private static WordTable CreatePaginationTable(WordDocument document, int rowCount) {
        document.Sections[0].ColumnCount = 2;
        document.Sections[0].ColumnsSpace = global::OfficeIMO.OfficeLength.Points(20);
        WordTable table = document.AddTable(rowCount, 1, WordTableStyle.TableGrid);
        // Keep frame capacity governed by the authored line heights; border
        // paint and cell margins have separate artifact regression coverage.
        table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Nil, global::OfficeIMO.OfficeLength.Points(0), OfficeIMO.OfficeColor.Black);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(4000) * 635)));
        for (int row = 0; row < rowCount; row++) {
            WordTableCell cell = table.Rows[row].Cells[0];
            cell.Width = WordTableWidth.Length(global::OfficeIMO.OfficeLength.Emus(checked((long)(4000) * 635)));
            cell.MarginTop = global::OfficeIMO.OfficeLength.Zero; cell.MarginBottom = global::OfficeIMO.OfficeLength.Zero; cell.MarginLeft = global::OfficeIMO.OfficeLength.Zero; cell.MarginRight = global::OfficeIMO.OfficeLength.Zero;
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = string.Join("\n", Enumerable.Range(1, 3).Select(line => $"Row{row + 1:D2}Line{line:D3}"));
            paragraph.FontFamily = "Arial"; paragraph.FontSize = global::OfficeIMO.OfficeLength.Points(12);
            paragraph.LineSpacing = global::OfficeIMO.Word.WordLineSpacing.Exact(global::OfficeIMO.OfficeLength.Emus(checked((long)(400) * 635)));
            paragraph.LineSpacingBefore = global::OfficeIMO.OfficeLength.Zero; paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Zero;
            paragraph.KeepLinesTogetherOverride = false; paragraph.KeepWithNextOverride = false;
            paragraph.AvoidWidowAndOrphanOverride = false;
        }
        return table;
    }

    private static WordPdfExportOptions CellPaginationOptions() => new() {
        IncludePageNumbers = false, PageSize = new PdfCore.PdfPageSize(OfficeIMO.OfficeLength.Points(500),OfficeIMO.OfficeLength.Points(400)), Margins = PdfCore.PdfPageMargins.Uniform(OfficeIMO.OfficeLength.Points(40))
    };
}
