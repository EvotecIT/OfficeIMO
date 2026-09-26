using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests;

public class PdfFloatingTableLayoutTests {
    [Theory]
    [InlineData(PdfTableVerticalAlignment.Center)]
    [InlineData(PdfTableVerticalAlignment.Bottom)]
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

}
