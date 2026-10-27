using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using W = DocumentFormat.OpenXml.Wordprocessing;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(100D, false, true, false, 2, false)]
    [InlineData(100D, true, true, false, 2, false)]
    [InlineData(160D, false, true, false, 2, false)]
    [InlineData(160D, true, true, false, 2, false)]
    [InlineData(200D, false, true, false, 2, false)]
    [InlineData(200D, true, true, false, 2, false)]
    [InlineData(100D, true, true, true, 2, false)]
    [InlineData(200D, false, true, true, 2, false)]
    [InlineData(80D, false, false, false, 1, false)]
    [InlineData(80D, false, true, false, 1, false)]
    [InlineData(80D, true, false, false, 1, false)]
    [InlineData(80D, true, true, false, 2, false)]
    [InlineData(100D, false, false, false, 1, false)]
    [InlineData(100D, false, false, true, 2, false)]
    [InlineData(160D, false, false, false, 2, false)]
    [InlineData(100D, false, true, false, 3, true)]
    [InlineData(100D, true, true, false, 3, true)]
    [InlineData(200D, false, true, false, 3, true)]
    [InlineData(200D, true, true, false, 3, true)]
    public void SaveAsPdf_ExactColumnLinesPreserveRunningStoriesAndPhysicalOverflow(
        double height, bool header, bool footer, bool widowControl, int pages, bool partial) {
        using WordDocument source = WordDocument.Create();
        source.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2013;
        source.Sections[0].ColumnCount = partial ? 1 : 2;
        source.Sections[0].ColumnsSpace = 240;
        source.Sections[0].Margins.HeaderDistance = source.Sections[0].Margins.FooterDistance = 720;
        if (header) SetColumnStory(source.HeaderDefaultOrCreate);
        if (footer) SetColumnStory(source.FooterDefaultOrCreate);
        if (partial) {
            WordParagraph preceding = source.AddParagraph("P");
            preceding.FontSize = 12;
            preceding.LineSpacingPoints = 12;
            preceding.LineSpacingRule = WordLineSpacingRule.Exact;
            preceding.LineSpacingBeforePoints = preceding.LineSpacingAfterPoints = 0;
            WordSection columns = source.AddSection(WordSectionBreakType.Continuous);
            columns.ColumnCount = 2;
            columns.ColumnsSpace = 240;
        }
        WordParagraph paragraph = source.AddParagraph();
        paragraph._paragraph.RemoveAllChildren<W.Run>();
        paragraph._paragraph.Append(TextLine("A", 12, true), TextLine("B", 12, false));
        paragraph.LineSpacingPoints = height;
        paragraph.LineSpacingRule = WordLineSpacingRule.Exact;
        paragraph.LineSpacingBeforePoints = paragraph.LineSpacingAfterPoints = 0;
        paragraph._paragraph.ParagraphProperties!.WidowControl = new W.WidowControl { Val = widowControl };
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        string before = document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml;
        byte[] bytes = document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(260),OfficeIMO.OfficeLength.Points(300)), Margins = new PdfPageMargins(OfficeIMO.OfficeLength.Points(30),OfficeIMO.OfficeLength.Points(72),OfficeIMO.OfficeLength.Points(30),OfficeIMO.OfficeLength.Points(72))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true });
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart!.Document.OuterXml);
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(pages, pdf.NumberOfPages);
        var first = Assert.Single(pdf.GetPage(partial ? 2 : 1).Letters, letter => letter.Value == "A");
        var second = Assert.Single(pdf.GetPage(pages).Letters, letter => letter.Value == "B");
        Assert.Equal(30D, first.StartBaseLine.X, 3);
        Assert.Equal(pages == 1 ? 136D : 30D, second.StartBaseLine.X, 3);
        // Independently exported Word controls retain the authored baseline on
        // initial and continuation pages, including ink below the body margin.
        double baseline = 300D - (header ? 132D : 72D) - height * .8D;
        Assert.Equal(baseline, first.StartBaseLine.Y, 3);
        Assert.Equal(baseline, second.StartBaseLine.Y, 3);
        if (partial) Assert.Equal("P", Assert.Single(pdf.GetPage(1).Letters).Value);
        Assert.Equal(pages == 1 ? 2 : 1, pdf.GetPage(partial ? 2 : 1).Letters.Count);
        if (pages > 1) Assert.Single(pdf.GetPage(pages).Letters);

    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SaveAsPdf_ExactColumnOverflowResetsBeforeOrdinaryColumnContinuation(bool balanced) {
        using WordDocument source = WordDocument.Create();
        source.CompatibilitySettings.CompatibilityMode = WordCompatibilityMode.Word2013;
        source.Sections[0].ColumnCount = 2;
        source.Sections[0].ColumnsSpace = 240;
        source.Sections[0].Margins.FooterDistance = 720;
        SetColumnStory(source.FooterDefaultOrCreate);
        WordParagraph tall = source.AddParagraph();
        tall._paragraph.Append(TextLine("A", 12, false));
        tall.LineSpacingPoints = 100;
        tall.LineSpacingRule = WordLineSpacingRule.Exact;
        tall.LineSpacingBeforePoints = tall.LineSpacingAfterPoints = 0;
        WordParagraph ordinary = source.AddParagraph();
        foreach (char letter in "BCDEFGHIJK") ordinary._paragraph.Append(TextLine(letter.ToString(), 12, letter != 'K'));
        ordinary.LineSpacingPoints = 12;
        ordinary.LineSpacingRule = WordLineSpacingRule.Exact;
        ordinary.LineSpacingBeforePoints = ordinary.LineSpacingAfterPoints = 0;
        ordinary._paragraph.ParagraphProperties!.WidowControl = new W.WidowControl { Val = false };
        if (balanced) {
            source.AddSection(WordSectionBreakType.Continuous).ColumnCount = 1;
            WordParagraph after = source.AddParagraph();
            after._paragraph.Append(TextLine("Z", 12, false));
            after.LineSpacingPoints = 12;
            after.LineSpacingRule = WordLineSpacingRule.Exact;
            after.LineSpacingBeforePoints = after.LineSpacingAfterPoints = 0;
        }
        using WordDocument document = WordDocument.Load(new MemoryStream(source.ToBytes()));
        using var pdf = PdfPigDocument.Open(document.ToPdf(new WordPdfExportOptions {
            IncludePageNumbers = false, ResourcePolicy = PdfResourcePolicy.CreatePortableDeterministic(),
            PageSize = new OfficeIMO.Pdf.PdfPageSize(OfficeIMO.OfficeLength.Points(260),OfficeIMO.OfficeLength.Points(300)), Margins = new PdfPageMargins(OfficeIMO.OfficeLength.Points(30),OfficeIMO.OfficeLength.Points(72),OfficeIMO.OfficeLength.Points(30),OfficeIMO.OfficeLength.Points(72))
        }).ToBytes(new OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal("A", Assert.Single(pdf.GetPage(1).Letters).Value);
        var letters = pdf.GetPage(2).Letters;
        Assert.Equal(balanced ? "BCDEFGHIJKZ" : "BCDEFGHIJK", string.Concat(letters.Select(letter => letter.Value)));
        int split = balanced ? 5 : 8;
        Assert.All(letters.Take(split), letter => Assert.Equal(30D, letter.StartBaseLine.X, 3));
        Assert.All(letters.Skip(split).Take(10 - split), letter => Assert.Equal(136D, letter.StartBaseLine.X, 3));
        Assert.Equal(letters[0].StartBaseLine.Y, letters[split].StartBaseLine.Y, 3);
        if (balanced) {
            Assert.Equal(30D, letters[10].StartBaseLine.X, 3);
            Assert.Equal(12D, letters[9].StartBaseLine.Y - letters[10].StartBaseLine.Y, 3);
        }
    }

    private static void SetColumnStory(WordHeaderFooter story) {
        for (int index = 0; index < 3; index++) {
            WordParagraph blank = story.AddParagraph();
            blank.LineSpacingPoints = 32;
            blank.LineSpacingRule = WordLineSpacingRule.Exact;
            blank.LineSpacingBeforePoints = blank.LineSpacingAfterPoints = 0;
        }
    }
}
