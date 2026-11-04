using System;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextExclusionFlowTests {
    private static readonly OfficeDrawingTextMetrics Metrics = new(
        (text, size, _, _) => (text?.Length ?? 0) * size / 2,
        (_, size, _, _) => new OfficeTextPaintBounds(-size * .8, size * .2),
        (text, size, _, _) => (0D, (text?.Length ?? 0) * size / 2));
    private static OfficeRichTextParagraph Paragraph(string text, double size = 10) =>
        new(new[] { new OfficeRichTextRun(text, size, OfficeColor.Black) });
    private static string Text(OfficeRichTextFlowFragment fragment) => string.Concat(fragment.Paragraph.Runs.Select(r => r.Text));
    private static readonly OfficeTextFlowRegion Area = new(0, 0, 100, 100);
    private static readonly OfficeTextFlowRegion[] Center = { new(40, 0, 20, 100) };

    [Fact]
    public void BothSidesReceiveSuccessiveSourceFragmentsOnTheSameRowBeforeAdvancing() {
        const string text = "aaa bbb ccc ddd eee fff";
        var flow = new OfficeRichTextFlow(new[] { Paragraph(text) }, _ => { });
        var fragments = flow.TakeAround(Area, Center, Metrics, () => { }, default);
        Assert.Equal(new[] { "aaa bbb ", "ccc ddd ", "eee fff" }, fragments.Select(Text));
        Assert.Equal(new[] { 0D, 60D, 0D }, fragments.Select(f => f.Region.X));
        Assert.Equal(new[] { 0D, 0D, 12D }, fragments.Select(f => Math.Round(f.Region.Y, 8)));
        Assert.Equal(text.Length, flow.CharacterPosition); Assert.False(flow.HasRemaining);
        foreach (var f in fragments) {
            var replay = OfficeDrawingTextLayout.CreateParagraphsWithMetrics(new[] { f.Paragraph }, f.Region.Width, f.Region.Height, Metrics);
            Assert.Single(replay.Lines); Assert.False(replay.Clipped);
        }
    }

    [Theory]
    [InlineData("A\r\nB", 2, 4)]
    [InlineData("A\r\n", 2, 3)]
    [InlineData("", 1, 0)]
    public void HardBreaksAndPendingEmptyLinesAdvanceRowsWithoutDuplicatingSource(string text, int count, int position) {
        var flow = new OfficeRichTextFlow(new[] { Paragraph(text) }, _ => { });
        var fragments = flow.TakeAround(Area, Center, Metrics, () => { }, default);
        Assert.Equal(count, fragments.Count); Assert.Equal(text, string.Concat(fragments.Select(Text)));
        Assert.All(fragments, f => Assert.Equal(0, f.Region.X));
        Assert.Equal(count == 2 ? 12D : 0, fragments[fragments.Count - 1].Region.Y, 8);
        Assert.Equal(position, flow.CharacterPosition); Assert.False(flow.HasRemaining);
    }

    [Fact]
    public void ParagraphSpacingAndLabelsApplyOnceAndDoNotStartTheNextParagraphInTheRightSlot() {
        var first = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("A", 10, OfficeColor.Black) },
            OfficeTextParagraphLabel.InBox(new OfficeRichTextRun("•", 10, OfficeColor.Black), 0, 5),
            margins: new OfficeTextPadding(0, 2, 0, 6));
        var second = new OfficeRichTextParagraph(new[] { new OfficeRichTextRun("B", 10, OfficeColor.Red) },
            margins: new OfficeTextPadding(0, 3, 0, 9));
        var flow = new OfficeRichTextFlow(new[] { first, second }, _ => { });
        var fragments = flow.TakeAround(Area, Center, Metrics, () => { }, default);
        Assert.Equal(new[] { "A", "B" }, fragments.Select(Text));
        Assert.Equal(new[] { (0D, 2D), (0D, 23D) }, fragments.Select(f => (f.Region.X, f.Region.Y)));
        Assert.NotNull(fragments[0].Paragraph.Label); Assert.Null(fragments[1].Paragraph.Label);
        Assert.All(fragments, f => Assert.Equal(0, f.Paragraph.Margins.Vertical));
        Assert.Equal(OfficeColor.Red, fragments[1].Paragraph.Runs[0].Color);
        Assert.Equal(3, flow.CharacterPosition); Assert.False(flow.HasRemaining);
        Assert.Equal(6, first.Margins.Bottom); Assert.Equal(9, second.Margins.Bottom);
    }

    [Fact]
    public void FullHeightMasksAreUnionedAndBlockedRowsResumeBelowTheirBottomWithoutConsumingText() {
        var masks = new[] { new OfficeTextFlowRegion(-10, 10, 70, 20), new OfficeTextFlowRegion(50, 10, 70, 20) };
        var flow = new OfficeRichTextFlow(new[] { Paragraph("A") }, _ => { });
        var fragment = Assert.Single(flow.TakeAround(Area, masks, Metrics, () => { }, default));
        Assert.Equal("A", Text(fragment)); Assert.Equal((0D, 30D), (fragment.Region.X, fragment.Region.Y));
        Assert.False(flow.HasRemaining); Assert.Equal(1, flow.CharacterPosition);
    }

    [Fact]
    public void AnUnselectedLargerFontDoesNotActivateAMaskBelowTheActualSmallFontRow() {
        var source = new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("AAAA ", 10, OfficeColor.Black), new OfficeRichTextRun("B", 20, OfficeColor.Red)
        });
        var flow = new OfficeRichTextFlow(new[] { source }, _ => { });
        var fragments = flow.TakeAround(Area, new[] {
            new OfficeTextFlowRegion(10, 0, 80, 100), new OfficeTextFlowRegion(0, 15, 100, 15)
        }, Metrics, () => { }, default);
        Assert.Equal(new[] { "AA", "AA " }, fragments.Take(2).Select(Text));
        Assert.All(fragments.Take(2), f => Assert.Equal(0, f.Region.Y));
        Assert.False(flow.HasRemaining); Assert.Equal(6, flow.CharacterPosition);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LongParagraphFlowMeasuresBoundedLookaheadAndRetainsOverflowUnderTheExistingBudget(bool tabFieldBeforeHardBreak) {
        const int limit = 4 * 1024 * 1024;
        string prefix = tabFieldBeforeHardBreak ? new string('A', 120) + "\tB\n" : string.Empty;
        string source = prefix + string.Concat(Enumerable.Repeat("alpha beta ", 9000));
        int measured = 0;
        var paragraph = Paragraph(source).WithTabStops(tabFieldBeforeHardBreak
            ? new OfficeTextTabStops(Array.Empty<OfficeTextTabStop>()) : null);
        var flow = new OfficeRichTextFlow(new[] { paragraph }, count => {
            measured = checked(measured + count);
            if (measured > limit) throw new InvalidOperationException("layout work budget");
        });
        var area = tabFieldBeforeHardBreak ? new OfficeTextFlowRegion(0, 0, 100, 1000) : new OfficeTextFlowRegion(0, 0, 360, 504);
        var masks = tabFieldBeforeHardBreak ? new[] { new OfficeTextFlowRegion(5, 0, 90, 1000) } : new[] {
            new OfficeTextFlowRegion(108, 60, 84, 84), new OfficeTextFlowRegion(72, 192, 108, 108)
        };
        var fragments = flow.TakeAround(area, masks, Metrics, () => { }, default);
        Assert.NotEmpty(fragments); Assert.True(flow.HasRemaining);
        Assert.Equal(source.Substring(0, flow.CharacterPosition), string.Concat(fragments.Select(Text)));
        Assert.InRange(measured, flow.CharacterPosition, limit);
        Assert.True(flow.CharacterPosition > prefix.Length);
    }

    [Fact]
    public void UnfittedRowRetainsSourceForTheNextFrameAndHonorsCancellationAndBothBudgets() {
        var flow = new OfficeRichTextFlow(new[] { Paragraph("A") }, _ => { });
        Assert.Empty(flow.TakeAround(new OfficeTextFlowRegion(0, 0, 100, 11.99), Center, Metrics, () => { }, default));
        Assert.True(flow.HasRemaining); Assert.Equal(0, flow.CharacterPosition);
        Assert.Equal("A", Text(Assert.Single(flow.TakeAround(Area, Center, Metrics, () => { }, default))));
        using var cancelled = new CancellationTokenSource(); cancelled.Cancel();
        Assert.Throws<OperationCanceledException>(() => flow.TakeAround(Area, Center, Metrics, () => { }, cancelled.Token));
        var budgeted = new OfficeRichTextFlow(new[] { Paragraph("A") }, _ => throw new InvalidOperationException("measurement"));
        Assert.Throws<InvalidOperationException>(() => budgeted.TakeAround(Area, Center, Metrics, () => { }, default));
        Assert.Equal(0, budgeted.CharacterPosition);
        Assert.Throws<InvalidOperationException>(() => budgeted.TakeAround(Area, Center, Metrics,
            () => throw new InvalidOperationException("inspection"), default));
    }
}
