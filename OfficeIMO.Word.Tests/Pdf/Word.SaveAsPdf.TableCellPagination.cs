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
            WordToPdfOptions options = CellPaginationOptions();
            options.PdfOptions = new PdfCore.PdfOptions { DefaultTableStyle = new PdfCore.PdfTableStyle {
                MinimumBodyRowsOnFirstPage = minimumRows, MinimumBodyRowsOnLastPage = minimumRows,
                HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 0, SpacingBefore = 0, SpacingAfter = 0
            } };
            document.SaveAsPdf(target, options);
        }
        using var pdf = PdfPigDocument.Open(target);
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(FindWordStartX(pdf.GetPage(1), "Row05Line001"), firstFrame ? 39.9 : 259.9, firstFrame ? 40.1 : 260.1);
    }

    private static WordTable CreatePaginationTable(WordDocument document, int rowCount) {
        document.Sections[0].ColumnCount = 2;
        document.Sections[0].ColumnsSpace = 400;
        WordTable table = document.AddTable(rowCount, 1);
        table.LayoutMode = WordTableLayoutMode.Fixed;
        table.Width = 4000; table.WidthType = WordTableWidthUnit.Dxa;
        for (int row = 0; row < rowCount; row++) {
            WordTableCell cell = table.Rows[row].Cells[0];
            cell.Width = 4000; cell.WidthType = WordTableWidthUnit.Dxa;
            cell.MarginTopWidth = 0; cell.MarginBottomWidth = 0; cell.MarginLeftWidth = 0; cell.MarginRightWidth = 0;
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = string.Join("\n", Enumerable.Range(1, 3).Select(line => $"Row{row + 1:D2}Line{line:D3}"));
            paragraph.FontFamily = "Arial"; paragraph.FontSize = 12;
            paragraph.LineSpacing = 400; paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
            paragraph.LineSpacingBeforePoints = 0; paragraph.LineSpacingAfterPoints = 0;
            paragraph.KeepLinesTogetherOverride = false; paragraph.KeepWithNextOverride = false;
            paragraph.AvoidWidowAndOrphanOverride = false;
        }
        return table;
    }

    private static WordToPdfOptions CellPaginationOptions() => new() {
        IncludePageNumbers = false, PageSize = new PdfCore.PageSize(500, 400), Margins = PdfCore.PageMargins.Uniform(40)
    };
}
