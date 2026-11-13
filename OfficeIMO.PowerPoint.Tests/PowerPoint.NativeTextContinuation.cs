using System.Text;
using OfficeIMO.PowerPoint;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PowerPointNativeTextContinuationTests {
    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OverflowingNativeParagraphContinuesAtSentenceOrWordBreaksWithoutLosingText(bool sentences) {
        var source = new StringBuilder();
        for (int index = 1; index <= 70; index++) {
            source.Append(sentences
                ? $"Finding {index:D2} traces the control to its owner and evidence. "
                : $"item{index:D2} control owner evidence{(index % 5 == 0 ? "  " : " ")}");
        }
        string text = source.ToString().TrimEnd();

        using var document = PowerPointDocument.Create();
        using (var materialization = new PowerPointLayoutMaterialization(document)) {
            var body = new PowerPointLayoutNativeText(materialization.CreateTextBody());
            PowerPointParagraph paragraph = body.Box.AddParagraph();
            paragraph.ClearInlineContent();
            paragraph.AddText(text);
            body.Freeze();
            var plan = new PowerPointContentPlanSlide("Continuation", new PowerPointLayoutBlock[] {
                new PowerPointLayoutText(text) { NativeBody = body }
            }, PowerPointLayout.Auto);
            materialization.Complete(plan, PowerPointCompositionOptions.FromDesign(
                PowerPointDeckDesign.FromBrand("#008C95", "native-continuation")));
        }

        string[] pieces = document.Slides
            .Select(slide => Assert.Single(slide.TextBoxes, box => box.Text.Contains(sentences ? "Finding" : "item")).Text)
            .ToArray();
        Assert.True(pieces.Length > 1, "The paragraph must overflow one slide.");
        Assert.Equal(text, string.Concat(pieces));
        foreach (string piece in pieces.Take(pieces.Length - 1)) {
            Assert.True(char.IsWhiteSpace(piece[piece.Length - 1]), "A continuation piece must end at a break: '" + piece + "'.");
            if (sentences) Assert.EndsWith(". ", piece, StringComparison.Ordinal);
        }
    }
}
