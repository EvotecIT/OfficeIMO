using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public partial class Word {
    [Fact]
    public void NestedListMarkerKeepsTheLargestParentIndex() {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(int.MaxValue));
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot));
        list.Numbering.Levels[1].LevelText = "%1.%2.";
        list.AddItem("Largest parent");
        WordParagraph child = list.AddItem("Nested item", 1);
        Assert.Equal("2147483647.1.", WordDocumentTraversal.BuildListMarkers(document)[child].Marker);
        Assert.Equal(1, WordDocumentTraversal.BuildListIndices(document)[child].Index);
        Assert.Empty(document.ValidateDocument());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ListCountersAllowTheLargestIndexAndRejectTheFollowingItem(bool markers) {
        using WordDocument document = WordDocument.Create();
        WordList list = document.AddCustomList();
        list.Numbering.AddLevel(new WordListLevel(WordListLevelKind.DecimalDot).SetStartNumberingValue(int.MaxValue));
        WordParagraph first = list.AddItem("Last representable index");
        if (markers) Assert.Equal("2147483647.", WordDocumentTraversal.BuildListMarkers(document)[first].Marker);
        else Assert.Equal(int.MaxValue, WordDocumentTraversal.BuildListIndices(document)[first].Index);
        list.AddItem("Overflowing index");
        Assert.Empty(document.ValidateDocument());
        InvalidDataException error = Assert.Throws<InvalidDataException>(() => {
            if (markers) WordDocumentTraversal.BuildListMarkers(document);
            else WordDocumentTraversal.BuildListIndices(document);
        });
        Assert.Contains("counter", error.Message, StringComparison.OrdinalIgnoreCase);
    }
}
