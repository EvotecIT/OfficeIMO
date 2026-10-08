using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Theory]
    [InlineData(WordListLevelKind.UpperRomanDot)]
    [InlineData(WordListLevelKind.LowerRomanDot)]
    public void InspectionRejectsExcessiveRomanListExpansion(WordListLevelKind kind) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(kind));
        list.Numbering.Levels[0].StartNumberingValue = int.MaxValue;
        list.AddItem("Large Roman start");
        string before = document._wordprocessingDocument.MainDocumentPart!.NumberingDefinitionsPart!.Numbering!.OuterXml;

        Assert.Contains("list marker", Assert.Throws<InvalidDataException>(() => document.CreateInspectionSnapshot()).Message);
        Assert.Contains("list marker", Assert.Throws<InvalidDataException>(() => WordDocumentTraversal.BuildListMarkers(document)).Message);
        Assert.Equal(before, document._wordprocessingDocument.MainDocumentPart.NumberingDefinitionsPart.Numbering.OuterXml);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ListMarkerLengthBoundsApplyToLiteralsAndExpandedPlaceholders(bool placeholders) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.UpperRoman));
        list.Numbering.Levels[0].StartNumberingValue = 1000000;
        list.Numbering.Levels[0].LevelText = placeholders ? "%1%1%1%1%1" : new string('x', 4097);
        list.AddItem("Long marker");

        Assert.Throws<InvalidDataException>(() => document.CreateInspectionSnapshot());
        list.Numbering.Levels[0].LevelText = new string('x', 4096);
        Assert.Equal(4096, Assert.Single(WordDocumentTraversal.BuildListMarkers(document)).Value.Marker.Length);
    }

    [Fact]
    public void InspectionBoundsTotalGeneratedListMarkerCharacters() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.Decimal));
        list.Numbering.Levels[0].LevelText = new string('x', 4096);
        for (int index = 0; index < 1024; index++) list.AddItem("Item " + index);
        Assert.Equal(1024, WordDocumentTraversal.BuildListMarkers(document).Count);
        list.AddItem("Excess item");
        Assert.Contains("character budget", Assert.Throws<InvalidDataException>(() => document.CreateInspectionSnapshot()).Message);
    }
}
