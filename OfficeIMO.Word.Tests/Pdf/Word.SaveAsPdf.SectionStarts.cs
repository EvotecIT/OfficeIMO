using DocumentFormat.OpenXml;
using W = DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    public static IEnumerable<object[]> SectionStartNumberingCases() {
        var cases = new[] {
            (WordSectionBreakType.OddPage, 1, 2, 0, 2), (WordSectionBreakType.OddPage, 2, 2, 0, 4),
            (WordSectionBreakType.EvenPage, 1, 2, 0, 3), (WordSectionBreakType.EvenPage, 2, 2, 0, 3),
            (WordSectionBreakType.NextPage, 1, 1, 2, 2), (WordSectionBreakType.NextPage, 2, 1, 2, 4),
            (WordSectionBreakType.NextPage, 1, 1, 3, 3), (WordSectionBreakType.NextPage, 2, 1, 3, 3),
            (WordSectionBreakType.OddPage, 1, 1, 2, 3), (WordSectionBreakType.OddPage, 2, 1, 2, 3),
            (WordSectionBreakType.EvenPage, 1, 1, 3, 2), (WordSectionBreakType.EvenPage, 2, 1, 3, 4)
        };
        foreach (var item in cases) foreach (WordFileFormat format in new[] { WordFileFormat.Docx, WordFileFormat.Doc })
            yield return new object[] { item.Item1, item.Item2, item.Item3, item.Item4, item.Item5, format };
    }

    [Theory]
    [MemberData(nameof(SectionStartNumberingCases))]
    public void SaveAsPdf_SectionStartUsesContinuingNumberBeforeRestart(
        WordSectionBreakType breakType, int firstPages, int firstNumber, int secondNumber, int expectedPages, WordFileFormat format) {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].Margins.Left = 800; document.Sections[0].Margins.Right = 1400;
        document.Settings.MirrorMargins = true;
        document.Sections[0].AddPageNumbering(firstNumber);
        for (int page = 1; page <= firstPages; page++) document.AddParagraph("First" + page).PageBreakBeforeOverride = page > 1;
        WordSection second = document.AddSection(breakType);
        if (secondNumber > 0) second.AddPageNumbering(secondNumber);
        second.AddParagraph("SecondSection");
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(format)));
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(expectedPages, pdf.NumberOfPages);
        Assert.Contains("SecondSection", pdf.GetPage(expectedPages).Text);
        int finalNumber = secondNumber > 0 ? secondNumber : firstNumber + expectedPages - 1;
        if (secondNumber > 0 && ((breakType == WordSectionBreakType.OddPage && finalNumber % 2 == 0) ||
            (breakType == WordSectionBreakType.EvenPage && finalNumber % 2 != 0))) finalNumber++;
        Assert.Equal(finalNumber % 2 == 0 ? 70D : 40D, FindWordStartX(pdf.GetPage(expectedPages), "SecondSection"), 2);
    }

    [Theory]
    [InlineData(WordSectionBreakType.NextPage, 1, 2, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.OddPage, 1, 3, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.EvenPage, 1, 2, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.NextPage, 2, 3, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.OddPage, 2, 3, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.EvenPage, 2, 4, WordFileFormat.Docx)]
    [InlineData(WordSectionBreakType.NextPage, 1, 2, WordFileFormat.Doc)]
    [InlineData(WordSectionBreakType.OddPage, 1, 3, WordFileFormat.Doc)]
    [InlineData(WordSectionBreakType.EvenPage, 1, 2, WordFileFormat.Doc)]
    [InlineData(WordSectionBreakType.NextPage, 2, 3, WordFileFormat.Doc)]
    [InlineData(WordSectionBreakType.OddPage, 2, 3, WordFileFormat.Doc)]
    [InlineData(WordSectionBreakType.EvenPage, 2, 4, WordFileFormat.Doc)]
    public void SaveAsPdf_SectionStartHonorsPhysicalParity(WordSectionBreakType breakType, int firstPages, int expectedPages, WordFileFormat format) {
        using WordDocument document = WordDocument.Create();
        for (int page = 1; page <= firstPages; page++) {
            WordParagraph paragraph = document.AddParagraph("First" + page);
            paragraph.PageBreakBeforeOverride = page > 1;
        }
        document.AddSection(breakType).AddParagraph("SecondSection");
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(format)));
        using PdfPigDocument pdf = PdfPigDocument.Open(loaded.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(expectedPages, pdf.NumberOfPages);
        for (int page = 1; page <= firstPages; page++) Assert.Contains("First" + page, pdf.GetPage(page).Text);
        Assert.Contains("SecondSection", pdf.GetPage(expectedPages).Text);
        if (expectedPages > firstPages + 1) Assert.True(string.IsNullOrWhiteSpace(pdf.GetPage(firstPages + 1).Text));
    }

    [Theory]
    [InlineData("OddPage", 3, WordFileFormat.Doc)]
    [InlineData("EvenPage", 4, WordFileFormat.Doc)]
    [InlineData("OddPage", 3, WordFileFormat.Docx)]
    [InlineData("EvenPage", 4, WordFileFormat.Docx)]
    public void SaveAsPdf_WordProducedSectionAndInlinePageBreaks(string type, int expectedPages, WordFileFormat format) {
        string extension = format == WordFileFormat.Doc ? "doc" : "docx";
        using WordDocument document = WordDocument.Load(GetFixtureDoc(Path.Combine("SectionStarts", $"word-section-{type}-2.{extension}")));
        Assert.Equal(2, document.Sections.Count);
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(expectedPages, pdf.NumberOfPages);
        Assert.Contains("First1", pdf.GetPage(1).Text);
        Assert.Contains("First2", pdf.GetPage(2).Text);
        Assert.Contains("SecondSection", pdf.GetPage(expectedPages).Text);
    }

    [Fact]
    public void SaveAsPdf_InlinePageBreakPreservesBothSidesAndRunFormatting() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        paragraph._paragraph.Append(new W.Run(new W.RunProperties(new W.Bold(), new W.FontSize { Val = "36" }),
            new W.Text("Before"), new W.Break { Type = W.BreakValues.Page }, new W.Text("After")));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Before", pdf.GetPage(1).Text);
        Assert.Contains("After", pdf.GetPage(2).Text);
        Assert.All(pdf.GetPage(1).Letters.Concat(pdf.GetPage(2).Letters), letter => Assert.Equal(18D, letter.FontSize, 2));
        Assert.All(pdf.GetPage(2).Letters, letter => Assert.Contains("Bold", letter.FontName, StringComparison.OrdinalIgnoreCase));
    }

    [Fact]
    public void SaveAsPdf_InlinePageBreakDoesNotRenderHiddenFieldInstructions() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        paragraph._paragraph.Append(
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Begin }),
            new W.Run(new W.FieldCode("PRIVATE"), new W.Break { Type = W.BreakValues.Page }),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.Separate }),
            new W.Run(new W.Text("Visible")),
            new W.Run(new W.FieldChar { FieldCharType = W.FieldCharValues.End }));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("Visible", pdf.GetPage(1).Text);
        Assert.DoesNotContain("PRIVATE", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_HiddenPageBreakRunDoesNotDiscardVisibleText() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        paragraph._paragraph.Append(new W.Run(new W.RunProperties(new W.Vanish()), new W.Break { Type = W.BreakValues.Page }),
            new W.Run(new W.Text("VisibleAfterHiddenBreak")));
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains("VisibleAfterHiddenBreak", pdf.GetPage(1).Text);
    }

    [Fact]
    public void SaveAsPdf_InlinePageBreaksRetainConsecutiveBlankPages() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph("First");
        paragraph.AddBreak(WordBreakType.Page).AddBreak(WordBreakType.Page).AddText("Third")
            .AddBreak(WordBreakType.Page).AddText("Fourth");
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(4, pdf.NumberOfPages);
        Assert.Contains("First", pdf.GetPage(1).Text);
        Assert.True(string.IsNullOrWhiteSpace(pdf.GetPage(2).Text));
        Assert.Contains("Third", pdf.GetPage(3).Text);
        Assert.Contains("Fourth", pdf.GetPage(4).Text);
    }

    [Fact]
    public void SaveAsPdf_InlinePageBreakAppliesFirstLineIndentOnlyOnce() {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].Margins.Left = 800;
        document.Sections[0].Margins.Right = 1400;
        document.Settings.MirrorMargins = true;
        WordParagraph paragraph = document.AddParagraph("Before");
        paragraph.IndentationBeforePoints = 30;
        paragraph.IndentationFirstLinePoints = 20;
        paragraph.AddBreak(WordBreakType.Page).AddText("After");
        using PdfPigDocument pdf = PdfPigDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false }));
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal(90D, FindWordStartX(pdf.GetPage(1), "Before"), 2);
        Assert.Equal(100D, FindWordStartX(pdf.GetPage(2), "After"), 2);
    }

    [Fact]
    public void SaveAsPdf_InlinePageBreakRetainsHyperlinkOnBothFragments() {
        using WordDocument document = WordDocument.Create();
        WordParagraph paragraph = document.AddParagraph();
        WordParagraph link = paragraph.AddHyperLink("Before", new Uri("https://example.com/section"));
        link._hyperlink!.Elements<W.Run>().Single().Append(new W.Break { Type = W.BreakValues.Page }, new W.Text("After"));
        byte[] bytes = document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false });
        using PdfPigDocument pdf = PdfPigDocument.Open(bytes);
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Contains("Before", pdf.GetPage(1).Text);
        Assert.Contains("After", pdf.GetPage(2).Text);
        var reader = OfficeIMO.Pdf.PdfReadDocument.Open(bytes);
        Assert.All(reader.Pages, page => Assert.Contains(page.GetLinkAnnotations(), annotation => annotation.Uri == "https://example.com/section"));
    }
}
