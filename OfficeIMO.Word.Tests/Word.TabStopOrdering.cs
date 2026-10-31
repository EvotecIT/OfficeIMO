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
        paragraph.LineSpacingAfter = global::OfficeIMO.OfficeLength.Points(12D);
        paragraph.AddTabStop(OfficeLength.Points(18), WordTabAlignment.Left);
        paragraph.AddTabStop(OfficeLength.Points(72), WordTabAlignment.Decimal);
        Assert.Empty(document.ValidateDocument());
        using var output = new MemoryStream();
        document.Save(output); output.Position = 0;
        using var reopened = WordDocument.Load(output);
        var restored = Assert.Single(reopened.Paragraphs);
        Assert.Equal(2, restored.TabStops.Count);
        restored.ClearTabStops().AddTabStop(OfficeLength.Points(36), WordTabAlignment.Right);
        Assert.Equal(OfficeLength.Points(36), Assert.Single(restored.TabStops).Position);
        Assert.Equal(12d, restored.LineSpacingAfter?.ToPoints());
        Assert.Equal(WordParagraphAlignment.Center, restored.ParagraphAlignment);
        Assert.Empty(reopened.ValidateDocument());
    }
}
