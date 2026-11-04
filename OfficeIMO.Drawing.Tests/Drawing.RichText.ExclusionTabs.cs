using System;
using System.Linq;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextExclusionTypographyTests {
    [Fact]
    public void EmptyHardBreakKeepsTypingFontGeometryFromBeyondTheLookaheadWindow() {
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), fonts: new OfficeFontFaceCollection()
            .Add("Publisher fixture", ManagedTextShapingTestAssets.CreatePublisherMetricsFont("dlig", "latn", false)));
        var metrics = canvas.CreateDrawingTextMetrics();
        var source = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("\n" + new string('A', 500), 10, OfficeColor.Black, fontFamily: "Publisher fixture"),
            new OfficeRichTextRun("B", 20, OfficeColor.Red, fontFamily: "Publisher fixture")
        }) { UsesPublisherFontMetrics = true };
        var flow = new OfficeRichTextFlow(new[] { source }, _ => { });
        var first = flow.TakeAround(new OfficeTextFlowRegion(0, 0, 100, 20),
            new[] { new OfficeTextFlowRegion(40, 0, 20, 100) }, metrics, () => { }, default);
        var fragment = Assert.Single(first);
        Assert.Equal("\n", string.Concat(fragment.Paragraph.Runs.Select(r => r.Text)));
        Assert.Equal(20D, fragment.Region.Height, 8);
        var replay = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { fragment.Paragraph },
            fragment.Region.Width, fragment.Region.Height, metrics);
        Assert.Equal(20D, Assert.Single(replay.Lines).ContentHeight!.Value, 8);
        Assert.True(flow.HasRemaining); Assert.Equal(1, flow.CharacterPosition);
    }

    [Fact]
    public void LookaheadKeepsSplitCrLfGraphemesAndTheCompleteAlignedTabField() {
        var metrics = new OfficeDrawingTextMetrics((text, size, _, _) => (text?.Length ?? 0) * size / 2,
            (_, size, _, _) => new OfficeTextPaintBounds(-size * .8, size * .2),
            (text, size, _, _) => (0D, (text?.Length ?? 0) * size / 2));
        string prefix = new string('A', 127);
        var source = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun(prefix + "\r", 10, OfficeColor.Black),
            new OfficeRichTextRun("\nCafé\u0301🐈\t" + new string('B', 160) + "\tEnd", 10, OfficeColor.Red)
        }).WithTabStops(new OfficeTextTabStops(new[] { new OfficeTextTabStop(150, OfficeTextTabAlignment.Right) }));
        var flow = new OfficeRichTextFlow(new[] { source }, _ => { });
        var fragments = flow.TakeAround(new OfficeTextFlowRegion(0, 0, 100, 1000),
            new[] { new OfficeTextFlowRegion(40, 0, 20, 1000) }, metrics, () => { }, default);
        Assert.False(flow.HasRemaining);
        string expected = string.Concat(source.Runs.Select(r => r.Text));
        Assert.Equal(expected, string.Concat(fragments.Select(f => string.Concat(f.Paragraph.Runs.Select(r => r.Text)))));
        Assert.Equal(expected.Length, flow.CharacterPosition);
        foreach (var fragment in fragments) {
            string text = string.Concat(fragment.Paragraph.Runs.Select(r => r.Text));
            if (text.Length == 0) continue;
            Assert.False(char.IsLowSurrogate(text[0])); Assert.False(char.IsHighSurrogate(text[text.Length - 1]));
            Assert.False(text.EndsWith("\r", StringComparison.Ordinal));
        }
    }

    [Fact]
    public void RightIntervalRetainsFrameRelativeTabGridAndReplaysItsSelectedAdvance() {
        var metrics = new OfficeDrawingTextMetrics((text, size, _, _) => (text?.Length ?? 0) * size / 2,
            (_, size, _, _) => new OfficeTextPaintBounds(-size * .8, size * .2),
            (text, size, _, _) => (0D, (text?.Length ?? 0) * size / 2));
        var paragraph = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("aaa bbb C\tD", 10, OfficeColor.Black) })
            .WithTabStops(new OfficeTextTabStops(Array.Empty<OfficeTextTabStop>(), defaultInterval: 36));
        var flow = new OfficeRichTextFlow(new[] { paragraph }, _ => { });
        var fragments = flow.TakeAround(new OfficeTextFlowRegion(10, 0, 100, 100),
            new[] { new OfficeTextFlowRegion(50, 0, 20, 100) }, metrics, () => { }, default);
        Assert.Equal(2, fragments.Count);
        Assert.Equal("C\tD", string.Concat(fragments[1].Paragraph.Runs.Select(r => r.Text)));
        Assert.Equal(-60D, fragments[1].Paragraph.TabStops!.Origin);
        var line = Assert.Single(OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { fragments[1].Paragraph },
            fragments[1].Region.Width, fragments[1].Region.Height, metrics).Lines);
        // Relative to the frame, the right slot begins at 60. C ends at 65;
        // the next default tab is 72, so D starts at 12 within this slot.
        Assert.Equal(17D, line.Width, 8);
        Assert.Equal(0D, paragraph.TabStops!.Origin); Assert.False(flow.HasRemaining);
    }

    [Fact]
    public void MixedFontFragmentsShareTheMeasuredBaselineAndWaitForTheWholeRowToFit() {
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), fonts: new OfficeFontFaceCollection()
            .Add("Publisher fixture", ManagedTextShapingTestAssets.CreatePublisherMetricsFont("dlig", "latn", false)));
        var metrics = canvas.CreateDrawingTextMetrics();
        var source = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("AAAA ", 10, OfficeColor.Black, fontFamily: "Publisher fixture"),
            new OfficeRichTextRun("B", 20, OfficeColor.Red, bold: true, fontFamily: "Publisher fixture")
        }) { UsesPublisherFontMetrics = true };
        var area = new OfficeTextFlowRegion(0, 0, 100, 100);
        var masks = new[] { new OfficeTextFlowRegion(25, 0, 35, 100) };
        var flow = new OfficeRichTextFlow(new[] { source }, _ => { });
        var fragments = flow.TakeAround(area, masks, metrics, () => { }, default);
        Assert.Equal(2, fragments.Count);
        Assert.Equal(new[] { "AAAA ", "B" }, fragments.Select(f => string.Concat(f.Paragraph.Runs.Select(r => r.Text))));
        double Baseline(OfficeRichTextFlowFragment f) {
            var line = Assert.Single(OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { f.Paragraph },
                f.Region.Width, f.Region.Height, metrics).Lines);
            return OfficeTextBlockRenderer.ResolveRichTextRenderBaseline(line, f.Region.Y, line.LineHeight, true);
        }
        Assert.Equal(Baseline(fragments[0]), Baseline(fragments[1]), 8);
        Assert.True(fragments[0].Region.Y > fragments[1].Region.Y);
        Assert.True(fragments[1].Paragraph.Runs[0].Bold); Assert.Equal(OfficeColor.Red, fragments[1].Paragraph.Runs[0].Color);
        var retained = new OfficeRichTextFlow(new[] { source }, _ => { });
        Assert.Empty(retained.TakeAround(new OfficeTextFlowRegion(0, 0, 100, 19.99), masks, metrics, () => { }, default));
        Assert.True(retained.HasRemaining); Assert.Equal(0, retained.CharacterPosition);
    }
}
