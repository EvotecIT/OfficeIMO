using System.IO;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Tests;

public class WordTabStopOrderingTests {
    [Fact]
    public void TabsAddedAfterSpacingRemainValidAndCanBeReplaced() {
        using var document = WordDocument.Create();
        var paragraph = document.AddParagraph("A\tB");
        paragraph.ParagraphAlignment = WordParagraphAlignment.Center;
        paragraph.LineSpacingAfterPoints = 12;
        paragraph.AddTabStop(360, WordTabAlignment.Left);
        paragraph.AddTabStop(1440, WordTabAlignment.Decimal);
        Assert.Empty(document.ValidateDocument());
        using var output = new MemoryStream();
        document.Save(output); output.Position = 0;
        using var reopened = WordDocument.Load(output);
        var restored = Assert.Single(reopened.Paragraphs);
        Assert.Equal(2, restored.TabStops.Count);
        restored.ClearTabStops().AddTabStop(720, WordTabAlignment.Right);
        Assert.Equal(720, Assert.Single(restored.TabStops).Position);
        Assert.Equal(12d, restored.LineSpacingAfterPoints);
        Assert.Equal(WordParagraphAlignment.Center, restored.ParagraphAlignment);
        Assert.Empty(reopened.ValidateDocument());
    }
}
