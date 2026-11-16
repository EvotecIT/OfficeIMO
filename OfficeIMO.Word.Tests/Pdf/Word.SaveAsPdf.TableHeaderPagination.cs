using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfCore = OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void SaveAsPdf_RepeatedHeaderMovesWithTheFirstFragmentOfAnOversizedBodyRow() {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2013;
        SetHeaderFlowParagraph(document.AddParagraph("FillerMarker"), 110);
        WordTable table = CreateHeaderFlowTable(document, 3);
        table.Rows[1].Cells[0].Paragraphs[0].Text = string.Join("\n",
            Enumerable.Range(1, 16).Select(index => $"BodyLine{index:D2}Marker"));
        using var pdf = RenderHeaderFlow(document);

        Assert.DoesNotContain("Header0Marker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("Header0Marker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("BodyLine01Marker", pdf.GetPage(2).Text, StringComparison.Ordinal);
        Assert.Contains("BodyLine16Marker", string.Concat(pdf.GetPages().Select(page => page.Text)), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(WordCompatibilityMode.Word2013, true)]
    [InlineData(WordCompatibilityMode.Word2010, false)]
    public void SaveAsPdf_RepeatedHeaderSplitsTheFirstBodyRowOnTheCurrentPageInModernLayout(WordCompatibilityMode mode,
        bool splitsOnFirstPage) {
        using WordDocument document = WordDocument.Create();
        document.CompatibilitySettings.CompatibilityMode = mode;
        SetHeaderFlowParagraph(document.AddParagraph("FillerMarker"), 90);
        WordTable table = CreateHeaderFlowTable(document, 3);
        table.Rows[1].Cells[0].Paragraphs[0].Text = string.Join("\n",
            Enumerable.Range(1, 4).Select(index => $"BodyLine{index:D2}Marker"));
        using var pdf = RenderHeaderFlow(document);

        Assert.Equal(splitsOnFirstPage, pdf.GetPage(1).Text.Contains("Header0Marker"));
        Assert.Equal(splitsOnFirstPage, pdf.GetPage(1).Text.Contains("BodyLine01Marker"));
        Assert.DoesNotContain("BodyLine02Marker", pdf.GetPage(1).Text, StringComparison.Ordinal);
        Assert.Contains("BodyLine04Marker", pdf.GetPage(2).Text, StringComparison.Ordinal);
    }

    private static PdfPigDocument RenderHeaderFlow(WordDocument document) {
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(WordFileFormat.Docx)));
        return PdfPigDocument.Open(loaded.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false,
            PageSize = new PdfCore.PdfPageSize(OfficeLength.Points(400), OfficeLength.Points(220)),
            Margins = PdfCore.PdfPageMargins.Uniform(OfficeLength.Points(40))
        }).ToBytes());
    }

    private static WordTable CreateHeaderFlowTable(WordDocument document, int rows) {
        WordTable table = document.AddTable(rows, 1, WordTableStyle.TableNormal);
        table.Width = WordTableWidth.Percent(100);
        table.StyleDetails!.SetBordersForAllSides(WordBorderStyle.Nil, OfficeLength.Zero, OfficeColor.Black);
        for (int index = 0; index < rows; index++) {
            WordTableRow row = table.Rows[index];
            row.RepeatHeaderRowAtTheTopOfEachPage = index == 0;
            WordTableCell cell = row.Cells[0];
            cell.MarginTop = cell.MarginBottom = cell.MarginLeft = cell.MarginRight = OfficeLength.Zero;
            WordParagraph paragraph = cell.Paragraphs[0];
            paragraph.Text = index == 0 ? "Header0Marker" : $"Body{index}Marker";
            SetHeaderFlowParagraph(paragraph, 20);
        }
        return table;
    }

    private static void SetHeaderFlowParagraph(WordParagraph paragraph, int lineHeight) {
        paragraph.FontFamily = "Arial"; paragraph.FontSize = OfficeLength.Points(12);
        paragraph.LineSpacing = WordLineSpacing.Exact(OfficeLength.Points(lineHeight));
        paragraph.LineSpacingBefore = paragraph.LineSpacingAfter = OfficeLength.Zero;
        paragraph.AvoidWidowAndOrphanOverride = false;
        paragraph.KeepWithNextOverride = false; paragraph.KeepLinesTogetherOverride = false;
    }
}
