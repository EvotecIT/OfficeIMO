using System;
using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextTabTests {
    private static double Measure(string? text, double size, string? family, OfficeFontStyle style) => (text?.Length ?? 0) * size / 2;
    private static OfficeRichTextBlockLayout Layout(OfficeRichTextParagraph paragraph, double width = 200) {
        var drawing = new OfficeDrawing(width, 200).AddRichTextParagraphs(new[] { paragraph }, 0, 0, width, 200);
        return OfficeDrawingTextLayout.Create((OfficeDrawingRichText)drawing.Elements[0], width, 200, Measure);
    }
    private static double Start(OfficeRichTextLine line, string text) => line.OffsetX + line.Segments.TakeWhile(s => s.Text != text).Sum(s => s.Width);

    [Theory]
    [InlineData(OfficeTextTabAlignment.Left, "123.45", 50)]
    [InlineData(OfficeTextTabAlignment.Center, "123.45", 35)]
    [InlineData(OfficeTextTabAlignment.Right, "123.45", 20)]
    [InlineData(OfficeTextTabAlignment.Character, "123.45", 35)]
    [InlineData(OfficeTextTabAlignment.Character, "123456", 20)]
    public void AlignsFollowingFieldAtMeasuredStop(OfficeTextTabAlignment alignment, string value, double start) {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\t" + value, 10, OfficeColor.Black) })
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(50, alignment) }));
        var line = Assert.Single(Layout(paragraph).Lines);
        Assert.Equal(start, Start(line, value));
        Assert.Equal("A" + value, string.Concat(line.Segments.Select(s => s.Text)));
        var drawing = new OfficeDrawing(200, 200).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 200, 200);
        Assert.Equal("A\t" + value, ((OfficeDrawingRichText)drawing.Elements[0]).PlainText);
    }

    [Fact]
    public void StyledRunBoundariesDoNotResetStopsOrCharacterFieldMeasurement() {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("AA", 10, OfficeColor.Black), new OfficeRichTextRun("\t12", 10, OfficeColor.Red),
            new OfficeRichTextRun(".50", 20, OfficeColor.Blue, bold: true), new OfficeRichTextRun("\tTail", 10, OfficeColor.Black)
        }).WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(60, OfficeTextTabAlignment.Character), new OfficeTextTabStop(120) }));
        var line = Assert.Single(Layout(paragraph).Lines);
        Assert.Equal(60, Start(line, ".50")); Assert.Equal(120, Start(line, "Tail"));
        Assert.Equal(OfficeColor.Blue, line.Segments.Single(s => s.Text == ".50").Color);
    }

    [Fact]
    public void ConsecutiveAndLeadingTabsAdvanceToDistinctStopsAndBreaksResetPosition() {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("\t\tA\n\tB", 10, OfficeColor.Black) })
            .WithTabStops(new OfficeTextTabStops(Array.Empty<OfficeTextTabStop>(), 30));
        var lines = Layout(paragraph).Lines;
        Assert.Equal(2, lines.Count); Assert.Equal(60, Start(lines[0], "A")); Assert.Equal(30, Start(lines[1], "B"));
    }

    [Fact]
    public void MarginsIndentationAndOriginsAreMeasuredInOneParagraphCoordinateSpace() {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB", 10, OfficeColor.Black) },
            margins: new OfficeTextPadding(20, 0, 0, 0), indent: OfficeTextParagraphIndent.FirstLine(10))
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(60) }, origin: -20));
        Assert.Equal(60, Start(Assert.Single(Layout(paragraph).Lines), "B"));
    }

    [Fact]
    public void ALabelMovingTheFirstLineRetainsTheParagraphTabOrigin() {
        var label = OfficeTextParagraphLabel.AtPosition(new OfficeRichTextRun("1", 10, OfficeColor.Black), 0,
            followedBy: OfficeTextParagraphLabelFollowedBy.Space);
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB", 10, OfficeColor.Black) }, label,
            margins: new OfficeTextPadding(40, 0, 0, 0)).WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(50) }));
        Assert.Equal(90, Start(Assert.Single(Layout(paragraph).Lines), "B"));
    }

    [Fact]
    public void OverlappingAlignedStopConsumesNoGapAndFollowingTabUsesTheNextStop() {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("AA\t1234\tB", 10, OfficeColor.Black) })
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(20, OfficeTextTabAlignment.Right), new OfficeTextTabStop(60, OfficeTextTabAlignment.Right) }, 36));
        var line = Assert.Single(Layout(paragraph).Lines);
        Assert.Equal(10, Start(line, "1234")); Assert.Equal(55, Start(line, "B"));
    }

    [Fact]
    public void DefaultsResumeOnTheGridAfterTheLastExplicitStop() {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB\tC", 10, OfficeColor.Black) })
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(50) }, 36));
        var line = Assert.Single(Layout(paragraph).Lines);
        Assert.Equal(50, Start(line, "B")); Assert.Equal(72, Start(line, "C"));
    }

    [Fact]
    public void TabBeyondFrameIsBoundedAndFollowingTextRemainsAvailable() {
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("Before\tAfter", 10, OfficeColor.Black) })
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(500) }));
        var layout = Layout(paragraph, 80);
        Assert.True(layout.Clipped); Assert.Equal(2, layout.Lines.Count);
        Assert.Equal("BeforeAfter", string.Concat(layout.Lines.SelectMany(l => l.Segments).Select(s => s.Text)));
        Assert.All(layout.Lines, l => Assert.True(l.OffsetX + l.Width <= 80));
    }

    [Theory]
    [InlineData(OfficeTextTabAlignment.Right)]
    [InlineData(OfficeTextTabAlignment.Center)]
    [InlineData(OfficeTextTabAlignment.Character)]
    public void LongAlignedFieldsDoNotMeasureEveryGrowingPrefix(OfficeTextTabAlignment alignment) {
        string body = "a. " + string.Concat(Enumerable.Repeat("a ", 4000));
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("\t" + body, 10, OfficeColor.Black) })
            .WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(50, alignment) }));
        var drawing = new OfficeDrawing(80, 200).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 80, 200);
        long measuredCharacters = 0;
        var layout = OfficeDrawingTextLayout.Create((OfficeDrawingRichText)drawing.Elements[0], 80, 200,
            (text, size, family, style) => {
                measuredCharacters += text?.Length ?? 0;
                return Measure(text, size, family, style);
            });
        Assert.True(layout.Clipped);
        // Bound deterministic measurement work, not host-dependent timing or allocation.
        Assert.InRange(measuredCharacters, 1, body.Length * 50L);
        Assert.StartsWith("a.", string.Concat(layout.Lines[0].Segments.Select(s => s.Text)));
    }

    [Fact]
    public void AlignedFieldMeasurementPreservesKerningAcrossEquivalentStyledRuns() {
        var paragraph = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("\tA", 10, OfficeColor.Black), new OfficeRichTextRun("V", 10, OfficeColor.Black)
        }).WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(50, OfficeTextTabAlignment.Right) }));
        var drawing = new OfficeDrawing(80, 200).AddRichTextParagraphs(new[] { paragraph }, 0, 0, 80, 200);
        var layout = OfficeDrawingTextLayout.Create((OfficeDrawingRichText)drawing.Elements[0], 80, 200,
            (text, size, family, style) => text == "AV" ? 7 : Measure(text, size, family, style));
        var line = Assert.Single(layout.Lines);
        Assert.Equal(43, Start(line, "AV"));
        Assert.Equal(50, line.Width);
    }

    [Theory]
    [InlineData(OfficeTextBaseline.Superscript)]
    [InlineData(OfficeTextBaseline.Subscript)]
    public void TabAdvanceRetainsTheRunsBaselineForLineHeight(OfficeTextBaseline baseline) {
        var settings = new OfficeTextTabStops(new[] { new OfficeTextTabStop(50) });
        var tabbed = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB", 10, OfficeColor.Black, baseline: baseline) })
            .WithTabStops(settings);
        var untabbed = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("AB", 10, OfficeColor.Black, baseline: baseline) });
        Assert.Equal(Assert.Single(Layout(untabbed).Lines).LineHeight, Assert.Single(Layout(tabbed).Lines).LineHeight);
    }

    [Fact]
    public void TabSettingsAreImmutableAndSurviveSceneCloneAndTint() {
        var stops = new[] { new OfficeTextTabStop(50) };
        var original = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A\tB", 10, OfficeColor.Red) });
        var settings = new OfficeTextTabStops(stops); var updated = original.WithTabStops(settings); stops[0] = new OfficeTextTabStop(70);
        Assert.Null(original.TabStops); Assert.Equal(50, updated.TabStops!.Stops[0].Position);
        var drawing = new OfficeDrawing(200, 200).AddRichTextParagraphs(new[] { updated }, 0, 0, 200, 200);
        var clone = drawing.Clone(); clone.ApplyColorTint(OfficeColor.Blue);
        var clonedParagraph = ((OfficeDrawingRichText)clone.Elements[0]).Paragraphs[0];
        Assert.Equal(50, clonedParagraph.TabStops!.Stops[0].Position);
        Assert.Equal(50, Start(Assert.Single(Layout(clonedParagraph).Lines), "B"));
        Assert.Equal(OfficeColor.Red, updated.Runs[0].Color);
        Assert.Throws<ArgumentException>(() => new OfficeTextTabStops(new[] { new OfficeTextTabStop(10), new OfficeTextTabStop(10) }));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeTextTabStops(stops, 0));
        Assert.Throws<ArgumentException>(() => new OfficeTextTabStop(10, character: "ab"));
    }
}
