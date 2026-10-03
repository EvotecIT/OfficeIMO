using OfficeIMO.Drawing;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("OddPage", 3, WordFileFormat.Doc)]
    [InlineData("EvenPage", 4, WordFileFormat.Doc)]
    [InlineData("OddPage", 3, WordFileFormat.Docx)]
    [InlineData("EvenPage", 4, WordFileFormat.Docx)]
    public void ImageExport_WordProducedSectionStartsDoNotDuplicateCachedBreaks(string type, int expectedPages, WordFileFormat format) {
        string extension = format == WordFileFormat.Doc ? "doc" : "docx";
        using WordDocument document = WordDocument.Load(GetFixtureDoc(Path.Combine("SectionStarts", $"word-section-{type}-2.{extension}")));
        IReadOnlyList<WordDocumentVisualSnapshot> pages = document.CreateVisualSnapshots();
        Assert.Equal(expectedPages, document.GetEstimatedPageCount());
        Assert.Equal(expectedPages, pages.Count);
        Assert.Contains(pages[0].Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "First1");
        Assert.Contains(pages[1].Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "First2");
        Assert.Contains(pages[expectedPages - 1].Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "SecondSection");
        if (expectedPages == 4) Assert.Empty(pages[2].Drawing.Elements.OfType<OfficeDrawingText>());
        Assert.All(pages, page => Assert.DoesNotContain(page.Diagnostics, diagnostic => diagnostic.Code == "unsupported-word-page-index"));
    }

    [Theory]
    [MemberData(nameof(SectionStartNumberingCases))]
    public void ImageExport_SectionStartUsesContinuingNumberBeforeRestart(
        WordSectionBreakType breakType, int firstPages, int firstNumber, int secondNumber, int expectedPages, WordFileFormat format) {
        using WordDocument document = WordDocument.Create();
        document.Sections[0].AddPageNumbering(firstNumber);
        for (int page = 1; page <= firstPages; page++) document.AddParagraph("First" + page).PageBreakBeforeOverride = page > 1;
        WordSection second = document.AddSection(breakType);
        if (secondNumber > 0) second.AddPageNumbering(secondNumber);
        second.AddParagraph("SecondSection");
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes(format)));
        IReadOnlyList<WordDocumentVisualSnapshot> pages = loaded.CreateVisualSnapshots();
        Assert.Equal(expectedPages, pages.Count);
        Assert.Contains(pages[expectedPages - 1].Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "SecondSection");
        for (int page = 0; page < firstPages; page++) Assert.Contains(pages[page].Drawing.Elements.OfType<OfficeDrawingText>(), text => text.Text == "First" + (page + 1));
    }
}
