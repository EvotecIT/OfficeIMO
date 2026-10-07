using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData("continued")]
    [InlineData("child")]
    [InlineData("parent")]
    public void NestedListCounterRestartsApplyOnceAndResetToTheAbstractStart(string restart) {
        using WordDocument document = CreateListInstanceCounterControl(out WordParagraph[] items, out Numbering numbering, out NumberingInstance second, 6);
        Level child = numbering.Elements<AbstractNum>().Single().Elements<Level>().Single(level => level.LevelIndex!.Value == 1);
        child.StartNumberingValue!.Val = 1;
        child.LevelText!.Val = "%1.%2.";
        child.NumberingFormat!.Val = NumberFormatValues.LowerLetter;
        if (restart == "child") second.Append(new LevelOverride(new StartOverrideNumberingValue { Val = 5 }) { LevelIndex = 1 });
        if (restart == "parent") second.Append(new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 });
        int[] ids = { 1, 1, 2, 1, 1, 2 }, levels = { 0, 1, 1, 1, 0, 1 };
        for (int index = 0; index < items.Length; index++) {
            NumberingProperties properties = items[index]._paragraph.ParagraphProperties!.NumberingProperties!;
            properties.NumberingId!.Val = ids[index]; properties.NumberingLevelReference!.Val = levels[index];
        }
        string[] expected = { "12.", "12.a.", restart == "child" ? "12.e." : "12.b.", restart == "child" ? "12.f." : "12.c.", "13.", "13.a." };
        int[] expectedIndices = { 12, 1, restart == "child" ? 5 : 2, restart == "child" ? 6 : 3, 13, 1 };
        var markers = WordDocumentTraversal.BuildListMarkers(document);
        var indices = WordDocumentTraversal.BuildListIndices(document);
        string text = PdfReadDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false })).ExtractText();
        for (int index = 0; index < items.Length; index++) {
            Assert.Equal(expected[index], markers[items[index]].Marker);
            Assert.Equal(expectedIndices[index], indices[items[index]].Index);
            Assert.Matches(System.Text.RegularExpressions.Regex.Escape(expected[index]) + @"\s*" + items[index].Text, text);
        }
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData("same")]
    [InlineData("template")]
    [InlineData("start")]
    [InlineData("text")]
    [InlineData("format")]
    [InlineData("zero")]
    [InlineData("absent")]
    public void SharedNumberingIdentityUsesTheFirstDefinitionAndContinuesItsCounters(string variation) {
        using WordDocument document = CreateListInstanceCounterControl(out WordParagraph[] items, out Numbering numbering, out NumberingInstance second);
        AbstractNum original = numbering.Elements<AbstractNum>().Single();
        original.Nsid = new Nsid { Val = variation == "zero" ? "00000000" : "11111111" };
        if (variation == "absent") original.Nsid.Remove();
        var definition = (AbstractNum)original.CloneNode(true);
        definition.AbstractNumberId = 1;
        if (variation == "template") definition.TemplateCode = new TemplateCode { Val = "33333333" };
        Level level = definition.Elements<Level>().First();
        if (variation == "start") level.StartNumberingValue!.Val = 7;
        if (variation == "text") level.LevelText!.Val = "%1)";
        if (variation == "format") level.NumberingFormat!.Val = NumberFormatValues.LowerRoman;
        numbering.InsertBefore(definition, numbering.Elements<NumberingInstance>().First());
        second.AbstractNumId!.Val = 1;
        items[3]._paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val = 1;
        string before = numbering.OuterXml;
        AssertListInstanceCounterValues(document, items, variation == "absent" ? new[] { 12, 13, 12, 14 } : new[] { 12, 13, 14, 15 }, ".");
        Assert.Equal(before, numbering.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void ListInstancesShareCountersInStoryOrderAndApplyExplicitRestartsOnce(bool interleaved, bool restart) {
        using WordDocument document = CreateListInstanceCounterControl(out WordParagraph[] items, out Numbering numbering, out NumberingInstance second);
        if (interleaved) items[3]._paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val = 1;
        if (restart) second.Append(new LevelOverride(new StartOverrideNumberingValue { Val = 1 }) { LevelIndex = 0 });
        string before = numbering.OuterXml;
        AssertListInstanceCounterValues(document, items, restart ? new[] { 12, 13, 1, 2 } : new[] { 12, 13, 14, 15 }, ".");
        Assert.Equal(before, numbering.OuterXml);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FormattingOnlyListOverridesKeepTheAbstractStartAndSequence(bool firstInstanceHasOverride) {
        using WordDocument document = CreateListInstanceCounterControl(out WordParagraph[] items, out Numbering numbering, out NumberingInstance second);
        Level level = (Level)numbering.Elements<AbstractNum>().Single().Elements<Level>().First().CloneNode(true);
        level.StartNumberingValue!.Val = 5;
        level.LevelText!.Val = "%1)";
        second.Append(new LevelOverride(level) { LevelIndex = 0 });
        if (firstInstanceHasOverride) foreach (WordParagraph item in items) item._paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val = 2;
        Assert.Equal(12, WordDocumentTraversal.GetListInfo(items[2])!.Value.Start);
        AssertListInstanceCounterValues(document, items, new[] { 12, 13, 14, 15 }, ")", firstInstanceHasOverride ? 0 : 2);
        Assert.Empty(document.ValidateDocument());
    }

    [Fact]
    public void SeparateListDefinitionsKeepIndependentCountersWhenInstancesAreInterleaved() {
        using WordDocument document = CreateListInstanceCounterControl(out WordParagraph[] items, out Numbering numbering, out NumberingInstance second);
        var definition = (AbstractNum)numbering.Elements<AbstractNum>().Single().CloneNode(true);
        definition.AbstractNumberId = 1;
        definition.Nsid = new Nsid { Val = "AABBCCDD" };
        definition.TemplateCode = new TemplateCode { Val = "A1B2C3D4" };
        numbering.InsertBefore(definition, numbering.Elements<NumberingInstance>().First());
        second.AbstractNumId!.Val = 1;
        items[3]._paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val = 1;
        AssertListInstanceCounterValues(document, items, new[] { 12, 13, 12, 14 }, ".");
        Assert.Empty(document.ValidateDocument());
    }

    private static WordDocument CreateListInstanceCounterControl(out WordParagraph[] items, out Numbering numbering, out NumberingInstance second, int count = 4) {
        WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Numbered);
        list.Numbering.Levels[0].StartNumberingValue = 12;
        foreach (int index in Enumerable.Range(0, count)) list.AddItem("ITEM-" + (char)('A' + index));
        numbering = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!;
        NumberingInstance first = numbering.Elements<NumberingInstance>().Single();
        first.RemoveAllChildren<LevelOverride>();
        second = (NumberingInstance)first.CloneNode(true);
        second.NumberID = 2;
        numbering.Append(second);
        items = document.Paragraphs.Where(item => item.Text.StartsWith("ITEM-", StringComparison.Ordinal)).ToArray();
        foreach (WordParagraph item in items.Skip(2)) item._paragraph.ParagraphProperties!.NumberingProperties!.NumberingId!.Val = 2;
        return document;
    }

    private static void AssertListInstanceCounterValues(WordDocument document, WordParagraph[] items, int[] expected, string suffix, int suffixStart = 0) {
        var indices = WordDocumentTraversal.BuildListIndices(document);
        var markers = WordDocumentTraversal.BuildListMarkers(document);
        string text = PdfReadDocument.Open(document.ToPdfBytes(new WordToPdfOptions { IncludePageNumbers = false })).ExtractText();
        for (int index = 0; index < items.Length; index++) {
            string marker = expected[index] + (index >= suffixStart ? suffix : ".");
            Assert.Equal(expected[index], indices[items[index]].Index);
            Assert.Equal(marker, markers[items[index]].Marker);
            Assert.Matches(System.Text.RegularExpressions.Regex.Escape(marker) + @"\s*" + items[index].Text, text);
        }
    }
}
