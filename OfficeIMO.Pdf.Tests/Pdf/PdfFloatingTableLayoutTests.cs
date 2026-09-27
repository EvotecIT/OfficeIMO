using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public class PdfFloatingTableLayoutTests {
    [Fact]
    public void FloatingTableKeepWithNextTransfersFullWidthFloatAndClearedLineTogether() {
        var style = Floating(320, 70); style.KeepWithNext = true;
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(240)).Canvas(c => c.Text("before", 40, 10, 100, 20)).Spacer(80)
            .Table(new[] { new[] { "floating" } }, style: style)
            .Paragraph(p => p.Text("below"), style: new PdfParagraphStyle { SpacingAfter = 0 }).ToBytes());
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "floating");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "floating");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "below");
    }

    [Fact]
    public void OuterKeptFlowAccountsForNestedKeptFlowClearance() {
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(240)).Canvas(c => c.Text("before", 40, 10, 100, 20)).Spacer(60).Flow(flow => {
            flow.Table(new[] { new[] { "floating" } }, style: Floating(100, 70));
            flow.Flow(nested => nested.Paragraph(p => p.Text("below"), style: new PdfParagraphStyle { LineHeight = 3.75, SpacingAfter = 0 }),
                new PdfFlowOptions { KeepTogether = true });
        }, new PdfFlowOptions { KeepTogether = true }).ToBytes());
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "floating");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "floating");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "below");
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void TrailingFloatingBookmarksStayAtExplicitPageAndSectionBoundaries(int boundary) {
        var document = PdfDocument.Create(Options(240));
        if (boundary == 0) document.Compose(builder => builder.Page(page => page.Content(content => content
            .Table(new[] { new[] { "floating" } }, style: Floating()).Bookmark("target"))));
        else document.Table(new[] { new[] { "floating" } }, style: Floating()).Bookmark("target")
            .Section("next", _ => { }, new PdfSectionOptions { StartOnNewPage = true });
        var destination = Assert.Single(PdfInspector.Inspect(document.Paragraph(p => p.Text("later")).ToBytes()).NamedDestinations.Where(d => d.Name == "target"));
        Assert.Equal(1, destination.PageNumber);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void TransparentNestedFlowsPreserveFloatingGeometry(int kind) {
        void Paint(PdfContentBuilder content) => content.Table(new[] { new[] { "floating" } }, style: Floating(100, 70));
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(240)).Spacer(60).Flow(flow => {
            if (kind == 0) flow.Semantic(PdfSemanticRole.Section, Paint);
            else if (kind == 1) flow.Layer("float", Paint);
            else flow.Flow(Paint);
            flow.Paragraph(p => p.Text("alongside"), style: new PdfParagraphStyle { SpacingAfter = 0 });
        }, new PdfFlowOptions { KeepTogether = true, OverflowBehavior = PdfFlowOverflowBehavior.Skip }).ToBytes());
        Assert.Contains(pdf.GetPage(1).GetWords(), w => w.Text == "alongside");
    }

    [Theory]
    [InlineData(PdfFlowOverflowBehavior.Skip)]
    [InlineData(PdfFlowOverflowBehavior.MoveToNextPage)]
    public void FloatingClearanceIsIncludedExactlyOnceInFlowMeasurement(PdfFlowOverflowBehavior overflow) {
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(240)).Spacer(60).Flow(flow => {
            flow.Table(new[] { new[] { "floating" } }, style: Floating(320, 70));
            flow.Paragraph(p => p.Text("below"), style: new PdfParagraphStyle { SpacingAfter = 0 });
        }, new PdfFlowOptions { OverflowBehavior = overflow }).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains(pdf.GetPage(1).GetWords(), w => w.Text == "below");
    }

    [Fact]
    public void FloatingTableOwnKeepWithNextMeasuresUnionWithFollowingText() {
        var style = Floating(100, 60); style.KeepWithNext = true;
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(240)).Spacer(100)
            .Table(new[] { new[] { "floating" } }, style: style)
            .Paragraph(p => p.Text("alongside"), style: new PdfParagraphStyle { SpacingAfter = 0 }).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains(pdf.GetPage(1).GetWords(), w => w.Text == "alongside");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PendingFloatBookmarkStaysBeforeExplicitPageBoundary(bool pageBlock) {
        var document = PdfDocument.Create(Options(240)).Table(new[] { new[] { "floating" } }, style: Floating()).Bookmark("target");
        if (pageBlock) document.Compose(builder => builder.Page(page => page.Content(content => content.Paragraph(p => p.Text("later")))));
        else document.PageBreak().Paragraph(p => p.Text("later"));
        var destination = Assert.Single(PdfInspector.Inspect(document.ToBytes()).NamedDestinations);
        Assert.Equal(1, destination.PageNumber);
        Assert.Equal(200, destination.DestinationTop!.Value, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PageAnchoredRowCannotUseHeightAbovePositiveOffset(bool deferred) {
        var style = Floating(120, 230);
        style.Position = new PdfTablePosition(verticalAnchor: PdfTableAnchor.Page, verticalOffset: 20);
        var document = PdfDocument.Create(Options(240));
        if (deferred) document.TableDeferred(() => new[] { new[] { "too tall" } }, batchSize: 1, style: style);
        else document.Table(new[] { new[] { "too tall" } }, style: style);
        Assert.Throws<System.ArgumentException>(() => document.ToBytes());
    }

    [Theory]
    [InlineData(PdfFlowOverflowBehavior.Skip)]
    [InlineData(PdfFlowOverflowBehavior.MoveToNextPage)]
    [InlineData(PdfFlowOverflowBehavior.Continue)]
    public void ConstrainedFlowMeasuresSideBySideFloatWithoutIgnoredSpacing(PdfFlowOverflowBehavior overflow) {
        var style = Floating(100, 70); style.SpacingBefore = 100; style.SpacingAfter = 100;
        var capture = new PdfLayoutPositionCapture();
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(240)).Spacer(60).Flow(flow => {
            flow.Table(new[] { new[] { "floating" } }, style: style);
            flow.Paragraph(p => p.Text("alongside"), style: new PdfParagraphStyle { SpacingAfter = 0 });
        }, new PdfFlowOptions { KeepTogether = true, OverflowBehavior = overflow }, capture: capture).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.Contains(pdf.GetPage(1).GetWords(), w => w.Text == "alongside");
        Assert.Contains(pdf.GetPage(1).GetWords(), w => w.Text == "floating");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CaptureIncludesContinuedTextBeforeFloatOnFinalPage(bool heading) {
        var capture = new PdfLayoutPositionCapture();
        string text = string.Join("\n", Enumerable.Range(0, heading ? 8 : 15).Select(i => "line" + i));
        byte[] bytes = PdfDocument.Create(Options(240)).Flow(flow => {
            if (heading) flow.H1(text);
            else flow.Paragraph(p => p.Text(text), style: new PdfParagraphStyle { LineHeight = 1, SpacingAfter = 0 });
            flow.Table(new[] { new[] { "floating" } }, style: Floating(120, 30));
        }, capture: capture).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.NumberOfPages > 1);
        var finalPage = pdf.GetPage(pdf.NumberOfPages);
        var firstContinuedWord = finalPage.GetWords().First(word => word.Text.StartsWith("line"));
        var region = capture.Regions.Single(item => item.PageNumber == pdf.NumberOfPages);
        Assert.Equal(320, region.Width, 3);
        Assert.True(region.Y + region.Height >= firstContinuedWord.BoundingBox.Top);
        var floating = finalPage.GetWords().Single(word => word.Text == "floating");
        Assert.True(region.Y <= floating.BoundingBox.Bottom);
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(2)]
    public void FloatingBookmarkResolvesAtFirstColumnOrCanvasContent(int kind) {
        var document = PdfDocument.Create(Options(240)).Table(new[] { new[] { "floating" } }, style: Floating(320, 80))
            .Bookmark("target");
        if (kind == 0) document.Columns(columns => columns.H1("first"));
        else if (kind == 1) document.Columns(columns => columns.Rectangle(60, 20));
        else document.Canvas(canvas => canvas.Text("first", 40, 130, 120, 20));
        byte[] bytes = document.PageBreak().Paragraph(p => p.Text("later")).ToBytes();
        var destination = Assert.Single(PdfInspector.Inspect(bytes).NamedDestinations);
        Assert.Equal(1, destination.PageNumber);
        Assert.InRange(destination.DestinationTop!.Value, 100, 130);
    }

    [Theory]
    [InlineData(false, false)]
    [InlineData(true, false)]
    [InlineData(false, true)]
    [InlineData(true, true)]
    public void MultiPageFloatOnlyCaptureExcludesUnusedFlowFrame(bool deferred, bool nested) {
        var capture = new PdfLayoutPositionCapture();
        var rows = new[] { new[] { "first" }, new[] { "second" } };
        var style = Floating(120, 90);
        System.Action<PdfContentBuilder> table = flow => {
            if (deferred) flow.TableDeferred(() => rows, batchSize: 1, style: style);
            else flow.Table(rows, style: style);
        };
        var document = PdfDocument.Create(Options(240));
        if (nested) document.Flow(flow => flow.Flow(table), capture: capture);
        else document.Flow(table, capture: capture);
        document.ToBytes();
        Assert.Equal(2, capture.Regions.Count);
        Assert.All(capture.Regions, region => {
            Assert.Equal(120, region.Width, 3);
            Assert.Equal(90, region.Height, 3);
            Assert.Equal(110, region.Y, 3);
        });
    }

    [Fact]
    public void AutomaticColumnsAssignSourceOrderAfterFloatClearance() {
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(240))
            .Table(new[] { new[] { "floating" } }, style: Floating(320, 80))
            .Columns(columns => {
                for (int index = 0; index < 12; index++) columns.Paragraph(p => p.Text("item" + index),
                    style: new PdfParagraphStyle { SpacingAfter = 0, LineHeight = 1 });
            }, new PdfMultiColumnOptions { ColumnCount = 2, BalanceLastPage = false }).ToBytes());
        var first = pdf.GetPage(1).GetWords().Where(w => w.Text.StartsWith("item")).Select(w => int.Parse(w.Text.Substring(4))).OrderBy(i => i).ToArray();
        Assert.Equal(Enumerable.Range(0, first.Length), first);
        Assert.Equal(12, pdf.GetPages().SelectMany(page => page.GetWords()).Count(w => w.Text.StartsWith("item")));
        Assert.All(pdf.GetPage(1).GetWords().Where(w => w.Text.StartsWith("item")), word => Assert.True(word.BoundingBox.Top < 121));
    }

    [Theory]
    [InlineData(PdfTableVerticalAlignment.Inside, true)]
    [InlineData(PdfTableVerticalAlignment.Outside, false)]
    public void VerticalInsideOutsideUsesPhysicalPageParity(PdfTableVerticalAlignment alignment, bool oddTop) {
        var style = Floating(120, 50);
        style.Position = new PdfTablePosition(verticalAnchor: PdfTableAnchor.Margin, verticalAlignment: alignment);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options())
            .Table(new[] { new[] { "odd" } }, style: style).PageBreak()
            .Table(new[] { new[] { "even" } }, style: style).ToBytes());
        Assert.Equal(oddTop, pdf.GetPage(1).GetWords().Single(w => w.Text == "odd").BoundingBox.Top > 300);
        Assert.Equal(!oddTop, pdf.GetPage(2).GetWords().Single(w => w.Text == "even").BoundingBox.Top > 300);
    }

    [Fact]
    public void OffsetFullWidthFloatDoesNotSplitKeptParagraph() {
        var style = Floating(320, 90); style.Position = new PdfTablePosition(verticalOffset: 25);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options(240)).Table(new[] { new[] { "floating" } }, style: style)
            .Paragraph(p => p.Text("first\nsecond\nthird\nfourth\nfifth"), style: new PdfParagraphStyle { KeepTogether = true }).ToBytes());
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), w => w.Text == "first");
        Assert.Equal(5, pdf.GetPage(2).GetWords().Count());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void StandaloneBookmarkFollowsActualParagraphPlacement(bool movePage) {
        var style = Floating(320, movePage ? 145 : 80);
        var document = PdfDocument.Create(Options(240)).Table(new[] { new[] { "floating" } }, style: style)
            .Bookmark("target");
        byte[] bytes = document.Paragraph(p => p.Text("targetword")).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var destination = Assert.Single(PdfInspector.Inspect(bytes).NamedDestinations);
        Assert.Equal(movePage ? 2 : 1, destination.PageNumber);
        var word = pdf.GetPage(destination.PageNumber!.Value).GetWords().Single(w => w.Text == "targetword");
        Assert.InRange(destination.DestinationTop!.Value - word.BoundingBox.Top, 0, 12);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void CapturedFlowIncludesFloatingTablePaintedBounds(bool nested) {
        var capture = new PdfLayoutPositionCapture();
        var style = Floating(120, 80); style.Position = new PdfTablePosition(verticalOffset: 50);
        var document = PdfDocument.Create(Options());
        if (nested) document.Flow(flow => flow.Flow(inner => inner.Table(new[] { new[] { "floating" } }, style: style)), capture: capture);
        else document.Flow(flow => flow.Table(new[] { new[] { "floating" } }, style: style), capture: capture);
        document.ToBytes();
        var region = Assert.Single(capture.Regions);
        Assert.Equal(80, region.Height, 3);
        Assert.Equal(120, region.Width, 3);
        Assert.Equal(330, region.Y, 3);
    }

    [Theory]
    [InlineData(PdfTableVerticalAlignment.Center)]
    [InlineData(PdfTableVerticalAlignment.Bottom)]
    [InlineData(PdfTableVerticalAlignment.Inside)]
    [InlineData(PdfTableVerticalAlignment.Outside)]
    public void DeferredFloatingTablesRejectAlignmentThatRequiresTotalHeight(PdfTableVerticalAlignment alignment) {
        var style = Floating(120, 30);
        style.Position = new PdfTablePosition(verticalAlignment: alignment);
        var document = PdfDocument.Create(Options()).TableDeferred(() => new[] { new[] { "one" }, new[] { "two" } }, batchSize: 1, style: style);
        Assert.Throws<System.ArgumentException>(() => document.ToBytes());
    }
    private static PdfOptions Options(double height = 500) => new() {
        PageWidth = 400, PageHeight = height, MarginLeft = 40, MarginRight = 40, MarginTop = 40, MarginBottom = 40
    };
    private static PdfTableStyle Floating(double width = 120, double height = 80) => new() {
        HeaderRowCount = 0, ColumnWidthPoints = new List<double?> { width }, MinRowHeight = height,
        Position = new PdfTablePosition()
    };

    [Fact]
    public void FloatingTableDoesNotSplitKeptParagraph() {
        byte[] bytes = PdfDocument.Create(Options(240)).Spacer(70)
            .Table(new[] { new[] { "floating" } }, style: Floating(200, 75))
            .Paragraph(paragraph => paragraph.Text(string.Join(" ", Enumerable.Range(1, 20).Select(index => "word" + index))),
                style: new PdfParagraphStyle { KeepTogether = true }).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text.StartsWith("word"));
        Assert.Equal(20, pdf.GetPage(2).GetWords().Count(word => word.Text.StartsWith("word")));
    }

    [Fact]
    public void WideInlineElementMovesBelowFloat() {
        byte[] bytes = PdfDocument.Create(Options())
            .Table(new[] { new[] { "floating" } }, style: Floating())
            .Paragraph(paragraph => paragraph.Inline(new PdfInlineBox(250, 20, background: PdfColor.Black)).Text("after"))
            .ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.Single(pdf.GetPage(1).GetWords(), word => word.Text == "after");
        Assert.True(pdf.GetPage(1).GetWords().Single(word => word.Text == "after").BoundingBox.Top < 380);
    }

    [Fact]
    public void DeferredBatchesShareOneAnchorAndRestoreFlow() {
        var style = Floating(120, 30);
        style.Position = new PdfTablePosition(PdfTableAnchor.Margin, PdfTableAnchor.Margin);
        byte[] bytes = PdfDocument.Create(Options())
            .TableDeferred(() => new[] { new[] { "one" }, new[] { "two" }, new[] { "three" } }, batchSize: 1, style: style)
            .Paragraph(paragraph => paragraph.Text("following")).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        var words = pdf.GetPage(1).GetWords().ToList();
        var one = words.Single(word => word.Text == "one");
        var two = words.Single(word => word.Text == "two");
        var three = words.Single(word => word.Text == "three");
        var following = words.Single(word => word.Text == "following");
        Assert.True(one.BoundingBox.Bottom - two.BoundingBox.Bottom > 25);
        Assert.True(two.BoundingBox.Bottom - three.BoundingBox.Bottom > 25);
        Assert.True(following.BoundingBox.Left >= 160);
        Assert.True(following.BoundingBox.Top > three.BoundingBox.Top);
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DefaultTablesAvoidPriorFloatingBounds(bool deferred) {
        var document = PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: Floating());
        if (deferred) document.TableDeferred(() => new[] { new[] { "normal" } }, batchSize: 1);
        else document.Table(new[] { new[] { "normal" } });
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.True(pdf.GetPage(1).GetWords().Single(word => word.Text == "normal").BoundingBox.Top < 380);
    }

    [Fact]
    public void FloatingCaptionKeepsFollowingTextOutsideCaptionBounds() {
        var style = Floating(); style.Caption = "caption"; style.CaptionSpacingAfter = 8;
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: style)
            .Paragraph(paragraph => paragraph.Text("following")).ToBytes());
        var words = pdf.GetPage(1).GetWords().ToList();
        Assert.True(words.Single(word => word.Text == "following").BoundingBox.Left >= 160);
    }

    [Fact]
    public void BottomPageAnchorStaysOnItsPage() {
        var style = Floating();
        style.Position = new PdfTablePosition(verticalAnchor: PdfTableAnchor.Page, verticalAlignment: PdfTableVerticalAlignment.Bottom);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "bottom" } }, style: style).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(pdf.GetPage(1).GetWords().Single(word => word.Text == "bottom").BoundingBox.Top, 0, 85);
    }

    [Fact]
    public void ForbiddenFloatingOverlapMovesSecondTableBelowFirst() {
        var style = Floating(); style.Position = new PdfTablePosition(allowOverlap: false);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "first" } }, style: style)
            .Table(new[] { new[] { "second" } }, style: style).ToBytes());
        var words = pdf.GetPage(1).GetWords().ToList();
        Assert.True(words.Single(word => word.Text == "first").BoundingBox.Bottom - words.Single(word => word.Text == "second").BoundingBox.Bottom >= 75);
    }

    [Fact]
    public void LaterWideInlineContentDoesNotDisplaceEarlierTextLines() {
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: Floating())
            .Paragraph(paragraph => paragraph.Text("early\nordinary\n")
                .Inline(new PdfInlineBox(250, 70, background: PdfColor.Black)).Text("late")).ToBytes());
        var words = pdf.GetPage(1).GetWords().ToList();
        var early = words.Single(word => word.Text == "early");
        var ordinary = words.Single(word => word.Text == "ordinary");
        Assert.True(early.BoundingBox.Left >= 160);
        Assert.InRange(early.BoundingBox.Bottom - ordinary.BoundingBox.Bottom, 10, 22);
        Assert.True(early.BoundingBox.Top > 430);
    }

    [Fact]
    public void MirroredPositionReversesTheInsideEdgeOnEvenPages() {
        var style = Floating(); style.Position = new PdfTablePosition(mirrorHorizontalOnEvenPages: true);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "odd" } }, style: style)
            .PageBreak().Table(new[] { new[] { "even" } }, style: style).ToBytes());
        Assert.InRange(pdf.GetPage(1).GetWords().Single(word => word.Text == "odd").BoundingBox.Left, 40, 50);
        Assert.InRange(pdf.GetPage(2).GetWords().Single(word => word.Text == "even").BoundingBox.Left, 240, 250);
    }

    [Fact]
    public void EarlierTableForbidsOverlapWithLaterDefaultTable() {
        var first = Floating(); first.Position = new PdfTablePosition(allowOverlap: false);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "first" } }, style: first)
            .Table(new[] { new[] { "second" } }, style: Floating()).ToBytes());
        var words = pdf.GetPage(1).GetWords().ToList();
        Assert.True(words.Single(word => word.Text == "first").BoundingBox.Bottom - words.Single(word => word.Text == "second").BoundingBox.Bottom >= 75);
    }

    [Theory]
    [InlineData(PdfTableVerticalAlignment.Bottom)]
    [InlineData(PdfTableVerticalAlignment.Center)]
    public void TallAnchoredTablesPaginateWithinPhysicalPage(PdfTableVerticalAlignment alignment) {
        var style = Floating(); style.Position = new PdfTablePosition(verticalAnchor: PdfTableAnchor.Page, verticalAlignment: alignment);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(Enumerable.Range(0, 8).Select(index => new[] { "row" + index }).ToArray(), style: style).ToBytes());
        Assert.True(pdf.NumberOfPages >= 2);
        var words = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).ToList();
        Assert.Equal(8, words.Count(word => word.Text.StartsWith("row")));
        Assert.All(words, word => Assert.InRange(word.BoundingBox.Top, 0, 500));
    }

    [Fact]
    public void AutomaticTablePageMoveRecalculatesInsideEdge() {
        var style = Floating(); style.KeepTogether = true;
        style.Position = new PdfTablePosition(mirrorHorizontalOnEvenPages: true);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Paragraph(paragraph => paragraph.Text("intro")).Spacer(350).Table(new[] { new[] { "even" } }, style: style).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.InRange(pdf.GetPage(2).GetWords().Single(word => word.Text == "even").BoundingBox.Left, 240, 250);
    }

    [Fact]
    public void DelimitedContinuationChunksRespectNarrowFloatingFrame() {
        var style = Floating(120, 220); style.Position = new PdfTablePosition(horizontalAlignment: PdfAlign.Right);
        const string token = "https://example.test/long_identifier_segment/another_long_identifier_segment/repeated_long_identifier_segment";
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: style)
            .Paragraph(paragraph => paragraph.Text(token)).ToBytes());
        Assert.All(pdf.GetPage(1).GetWords().Where(word => word.Text != "floating" && word.BoundingBox.Top > 240),
            word => Assert.True(word.BoundingBox.Right <= 241));
    }

    [Fact]
    public void BottomAlignedContinuationIncludesRepeatedHeaderHeight() {
        var style = Floating(); style.HeaderRowCount = 1; style.RepeatHeaderRowCount = 1;
        style.Position = new PdfTablePosition(verticalAnchor: PdfTableAnchor.Page, verticalAlignment: PdfTableVerticalAlignment.Bottom);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(Enumerable.Range(0, 8).Select(index => new[] { "row" + index }).ToArray(), style: style).ToBytes());
        Assert.True(pdf.NumberOfPages >= 2);
        var words = Enumerable.Range(1, pdf.NumberOfPages).SelectMany(page => pdf.GetPage(page).GetWords()).ToList();
        Assert.All(words, word => Assert.InRange(word.BoundingBox.Bottom, 0, 500));
        Assert.Single(words, word => word.Text == "row7");
    }

    [Fact]
    public void WordRechecksWidthWhenItsNewLineEntersFloat() {
        var style = Floating(200, 220); style.Position = new PdfTablePosition(horizontalAlignment: PdfAlign.Right, verticalOffset: 20);
        string pending = new string('W', 20);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: style)
            .Paragraph(paragraph => paragraph.Text("prefixprefixprefixprefixprefix " + pending)).ToBytes());
        var word = pdf.GetPage(1).GetWords().Single(word => word.Text == pending);
        Assert.True(word.BoundingBox.Top <= 221, word.BoundingBox.ToString());
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void DefaultPositionedStyleRetainsSharedFlowAnchor(bool deferred) {
        var options = Options(); options.DefaultTableStyle = Floating();
        var document = PdfDocument.Create(options).Table(new[] { new[] { "first" } });
        if (deferred) document.TableDeferred(() => new[] { new[] { "second" } }, batchSize: 1);
        else document.Table(new[] { new[] { "second" } });
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        var letters = pdf.GetPage(1).Letters;
        var first = letters.Single(letter => letter.Value == "f");
        var second = letters.Single(letter => letter.Value == "s" && letter.Location.X == first.Location.X);
        Assert.Equal(first.Location.Y, second.Location.Y, 2);
    }

    [Fact]
    public void FlowRuleAndFormFieldStayBelowFloat() {
        byte[] bytes = PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: Floating())
            .HR().TextField("field", value: "value").ToBytes();
        var widget = Assert.Single(PdfInspector.Inspect(bytes).GetFormWidgets("field"));
        Assert.True(widget.Y2 <= 380);
    }

    [Fact]
    public void FlowAnnotationAvoidsAnOffsetFloat() {
        var style = Floating(); style.Position = new PdfTablePosition(verticalOffset: 20);
        byte[] bytes = PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: style)
            .FreeTextAnnotation("note", 120, 40).ToBytes();
        Assert.True(Assert.Single(PdfInspector.Inspect(bytes).GetAnnotationsBySubtype("FreeText")).Y2 <= 360);
    }

    [Fact]
    public void AutomaticColumnsStayBelowFloat() {
        byte[] bytes = PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: Floating())
            .Columns(columns => columns.Paragraph(paragraph => paragraph.Text("column"))).ToBytes();
        using var pdf = PdfPigDocument.Open(bytes);
        Assert.True(pdf.GetPage(1).GetWords().Single(word => word.Text == "column").BoundingBox.Top <= 381);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ContinuedFloatIgnoresFlowSpacingAfter(bool deferred) {
        double LastBaseline(double spacing) {
            var style = Floating(); style.SpacingAfter = spacing;
            var rows = Enumerable.Range(0, 8).Select(index => new[] { "row" + index }).ToArray();
            var document = PdfDocument.Create(Options());
            if (deferred) document.TableDeferred(() => rows, batchSize: 3, style: style);
            else document.Table(rows, style: style);
            using var pdf = PdfPigDocument.Open(document.Paragraph(paragraph => paragraph.Text("after")).ToBytes());
            return pdf.GetPage(pdf.NumberOfPages).GetWords().Single(word => word.Text == "after").BoundingBox.Top;
        }
        Assert.Equal(LastBaseline(0), LastBaseline(50), 2);
    }

    [Theory]
    [InlineData(PdfTableAnchor.Margin, 44)]
    [InlineData(PdfTableAnchor.Flow, 174)]
    public void NestedTablesDistinguishPageMarginAndFlowAnchors(PdfTableAnchor anchor, double expected) {
        var table = Floating(); table.Position = new PdfTablePosition(horizontalAnchor: anchor);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Panel(panel => panel.Table(new[] { new[] { "nested" } }, style: table),
            new PdfPanelStyle { MaxWidth = 200, Align = PdfAlign.Right, PaddingX = 10, PaddingY = 0, KeepTogether = false }).ToBytes());
        Assert.Equal(expected, pdf.GetPage(1).GetWords().Single(word => word.Text == "nested").BoundingBox.Left, 1);
    }

    [Fact]
    public void RowColumnRejectsPositionedTableInsteadOfIgnoringIt() {
        var document = PdfDocument.Create(Options()).Compose(builder => builder.Page(page => page.Content(content =>
            content.Row(row => row.PercentColumn(100, column => column.Table(new[] { new[] { "nested" } }, style: Floating()))))));
        Assert.Throws<System.NotSupportedException>(() => document.ToBytes());
    }

    [Fact]
    public void FlowSpacingDoesNotChangeBottomAnchorOrForceExtraPage() {
        var style = Floating(); style.SpacingBefore = 50; style.SpacingAfter = 50; style.KeepTogether = true;
        style.Position = new PdfTablePosition(verticalAnchor: PdfTableAnchor.Page, verticalAlignment: PdfTableVerticalAlignment.Bottom);
        using var pdf = PdfPigDocument.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "bottom" } }, style: style).ToBytes());
        Assert.Equal(1, pdf.NumberOfPages);
        Assert.InRange(pdf.GetPage(1).GetWords().Single(word => word.Text == "bottom").BoundingBox.Top, 0, 85);
    }

}
