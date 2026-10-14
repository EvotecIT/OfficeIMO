using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextLabelTests {
    [Theory]
    [InlineData(OfficeTextAlignment.Left, 10)]
    [InlineData(OfficeTextAlignment.Center, 37.5)]
    [InlineData(OfficeTextAlignment.Right, 65)]
    public void NothingSeparatorRetainsConcatenatedRenderedTextAndAlignedBodyPosition(OfficeTextAlignment alignment, double bodyPosition) {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("123", 10, OfficeColor.Black) },
            OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("ID", 10, OfficeColor.Black), 0,
                followedBy: OfficeTextParagraphLabelFollowedBy.Nothing), alignment);
        var line = Assert.Single(Layout(paragraph).Lines);
        Assert.Equal("ID123", string.Concat(line.Segments.Select(s => s.Text)));
        Assert.Equal(bodyPosition, line.OffsetX + line.Segments.TakeWhile(s => s.Text != "123").Sum(s => s.Width));
    }

    [Theory]
    [InlineData(30, 0, 0)]
    [InlineData(2, 75, 0)]
    [InlineData(2, 65, 10)]
    public void OversizedLabelDoesNotPaintOutsideItsTextFrameOrMoveBody(int length, double position, double rightMargin) {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Body", 10, OfficeColor.Black) },
            OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun(new string('W', length), 10, OfficeColor.Black), position),
            margins: new OfficeTextPadding(20, 0, rightMargin, 0));
        var layout = Layout(paragraph);
        Assert.True(layout.Clipped);
        Assert.DoesNotContain(layout.Lines.SelectMany(l => l.Segments), s => s.Text.Contains('W'));
        Assert.Equal(20, Assert.Single(layout.Lines).OffsetX);
        var scene = new OfficeDrawing(200, 100).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 80, 100);
        Assert.DoesNotContain(new string('W', length), OfficeDrawingSvgExporter.ToSvg(scene));
        Assert.StartsWith(new string('W', length), ((OfficeDrawingRichText)scene.Elements[0]).PlainText);
    }
    private static double Measure(string? text, double size, string? family, OfficeFontStyle style) => (text?.Length ?? 0) * size / 2;
    private static OfficeRichTextBlockLayout Layout(OfficeRichTextParagraph paragraph, double width = 80) {
        var scene = new OfficeDrawing(width, 200).AddRichTextParagraphs(new[] { paragraph }, 0, 0, width, 200);
        return OfficeDrawingTextLayout.Create((OfficeDrawingRichText)scene.Elements[0], width, 200, Measure);
    }

    [Fact]
    public void LabelIsMeasuredAtRenderTimeAndOnlyOccursOnTheFirstWrappedLine() {
        var label = OfficeTextParagraphLabel.InBox(new OfficeRichTextRun("12.", 10, OfficeColor.Red), 0, 20, OfficeTextAlignment.Right, 4);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("one two three four five", 10, OfficeColor.Black) }, label, margins: new OfficeTextPadding(20, 0, 0, 0));
        var layout = Layout(paragraph);
        Assert.Equal("12.", layout.Lines[0].Segments[0].Text); Assert.Equal(OfficeColor.Red, layout.Lines[0].Segments[0].Color);
        Assert.Equal(1, layout.Lines[0].OffsetX); // The gap moves the label within its box before moving body text.
        Assert.All(layout.Lines.Skip(1), line => { Assert.Equal(20, line.OffsetX); Assert.DoesNotContain(line.Segments, segment => segment.Text == "12."); });
        var scene = new OfficeDrawing(80, 200).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 80, 200);
        Assert.StartsWith("12. one", ((OfficeDrawingRichText)scene.Elements[0]).PlainText);
        Assert.Contains("12.", OfficeDrawingSvgExporter.ToSvg(scene.Clone()));
    }

    [Theory]
    [InlineData(OfficeTextParagraphLabelFollowedBy.Space, 25)]
    [InlineData(OfficeTextParagraphLabelFollowedBy.Nothing, 20)]
    [InlineData(OfficeTextParagraphLabelFollowedBy.Position, 35)]
    public void LabelFollowingRuleDoesNotChangeContinuationIndentation(OfficeTextParagraphLabelFollowedBy following, double textStart) {
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("1.", 10, OfficeColor.Black), 10, followedBy: following, textPosition: 35);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("first\nsecond", 10, OfficeColor.Black) }, label, margins: new OfficeTextPadding(30, 0, 0, 0));
        var layout = Layout(paragraph);
        Assert.Equal(textStart, layout.Lines[0].OffsetX + layout.Lines[0].Segments.TakeWhile(segment => segment.Text != "first").Sum(segment => segment.Width));
        Assert.Equal(30, layout.Lines[1].OffsetX);
    }

    [Fact]
    public void EmptyBodyStillPaintsLabelAndLargeLabelSetsAutomaticLineHeight() {
        var label = OfficeTextParagraphLabel.InBox(new OfficeRichTextRun("•", 30, OfficeColor.Black), 0, 20);
        var paragraph = new OfficeRichTextParagraph(Array.Empty<OfficeRichTextRun>(), label, margins: new OfficeTextPadding(20, 0, 0, 0));
        var line = Assert.Single(Layout(paragraph).Lines); Assert.Equal("•", line.Segments[0].Text); Assert.True(line.LineHeight >= 30);
        Assert.Throws<ArgumentException>(() => OfficeTextParagraphLabel.InBox(new OfficeRichTextRun("bad\nlabel", 10, OfficeColor.Black), 0, 20));
    }

    [Fact]
    public void TintingAClonedSceneRetainsLabelLayoutWithoutChangingOriginal() {
        var label = OfficeTextParagraphLabel.InBox(new OfficeRichTextRun("3.", 12, OfficeColor.Red), 4, 24, OfficeTextAlignment.Right, 3);
        var original = new OfficeDrawing(100, 100).AddRichTextParagraphs(new[] { new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Body", 12, OfficeColor.Black) }, label) }, 0, 0, 100, 100);
        var clone = original.Clone(); clone.ApplyColorTint(OfficeColor.Blue);
        var tinted = ((OfficeDrawingRichText)clone.Elements[0]).Paragraphs[0].Label!;
        Assert.Equal(OfficeColor.Blue, tinted.Run.Color); Assert.Equal(4, tinted.Position); Assert.Equal(24, tinted.MinimumWidth); Assert.Equal(3, tinted.MinimumDistance);
        Assert.Equal(OfficeColor.Red, ((OfficeDrawingRichText)original.Elements[0]).Paragraphs[0].Label!.Run.Color);
        Assert.Contains("3.", OfficeDrawingSvgExporter.ToSvg(clone));
    }

    [Fact]
    public void LabelAndFollowingSpaceParticipateInSharedCharacterBudget() {
        var label = OfficeTextParagraphLabel.InBox(new OfficeRichTextRun("1.", 12, OfficeColor.Black), 0, 20);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(new string('a', 99998), 12, OfficeColor.Black) }, label);
        Assert.Throws<ArgumentException>(() => new OfficeDrawing(100, 100).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 100, 100));
        Assert.Throws<ArgumentException>(() => new OfficeRichTextParagraph(new[] { new OfficeRichTextRun(new string('a', 99999), 12, OfficeColor.Black) }, label));
    }
}
