using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using DocumentFormat.OpenXml.Wordprocessing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void FreshListLevelStartControlsSavedAndProjectedNumbering() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Numbered);
        list.Numbering.Levels[0].StartNumberingValue = 12;
        list.AddItem("First"); list.AddItem("Second");
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        WordParagraph first = loaded.Paragraphs.First(paragraph => paragraph.Text == "First");
        var main = loaded._wordprocessingDocument.MainDocumentPart!;
        int numberId = first._paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val!.Value;
        var numbering = main.NumberingDefinitionsPart!.Numbering!;
        NumberingInstance instance = numbering.Elements<NumberingInstance>().Single(item => item.NumberID!.Value == numberId);
        AbstractNum definition = numbering.Elements<AbstractNum>().Single(item => item.AbstractNumberId!.Value == instance.AbstractNumId!.Val!.Value);
        int savedStart = instance.Elements<LevelOverride>().FirstOrDefault(item => item.LevelIndex!.Value == 0)?.StartOverrideNumberingValue?.Val?.Value
            ?? definition.Elements<Level>().Single(item => item.LevelIndex!.Value == 0).StartNumberingValue!.Val!.Value;
        Assert.Equal(12, savedStart);
        AssertListMarkers(loaded, "12.", "13.");
        Assert.Empty(loaded.ValidateDocument());
    }

    [Theory]
    [InlineData(1)]
    [InlineData(9)]
    public void AuthoredInstanceStartOneOverridesAbstractStartRegardlessOfOtherLevels(int overrideCount) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Numbered);
        list.Numbering.Levels[0].StartNumberingValue = 12;
        list.AddItem("First"); list.AddItem("Second");
        var numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        NumberingInstance instance = numbering.Elements<NumberingInstance>().Single();
        instance.RemoveAllChildren<LevelOverride>();
        for (int level = 0; level < overrideCount; level++)
            instance.Append(new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = level });
        using WordDocument loaded = WordDocument.Load(new MemoryStream(document.ToBytes()));
        string before = loaded._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.OuterXml;
        AssertListMarkers(loaded, "1.", "2.");
        Assert.Equal(1, WordDocumentTraversal.GetListInfo(loaded.Paragraphs.First(paragraph => paragraph.Text == "First"))!.Value.Start);
        Assert.Equal(before, loaded._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.OuterXml);
        Assert.Empty(loaded.ValidateDocument());
    }

    private static void AssertListMarkers(WordDocument document, string first, string second) {
        var markers = WordDocumentTraversal.BuildListMarkers(document);
        Assert.Equal(first, markers[document.Paragraphs.First(paragraph => paragraph.Text == "First")].Marker);
        Assert.Equal(second, markers[document.Paragraphs.First(paragraph => paragraph.Text == "Second")].Marker);
        string pdfText = PdfReadDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false })).ExtractText();
        Assert.Matches(System.Text.RegularExpressions.Regex.Escape(first) + @"\s*First", pdfText);
        Assert.Matches(System.Text.RegularExpressions.Regex.Escape(second) + @"\s*Second", pdfText);
    }
}
