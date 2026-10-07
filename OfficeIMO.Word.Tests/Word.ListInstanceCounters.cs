using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Pdf;
using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
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

    private static WordDocument CreateListInstanceCounterControl(out WordParagraph[] items, out Numbering numbering, out NumberingInstance second) {
        WordDocument document = WordDocument.Create();
        WordList list = document.AddList(WordListStyle.Numbered);
        list.Numbering.Levels[0].StartNumberingValue = 12;
        foreach (string text in new[] { "ITEM-A", "ITEM-B", "ITEM-C", "ITEM-D" }) list.AddItem(text);
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
