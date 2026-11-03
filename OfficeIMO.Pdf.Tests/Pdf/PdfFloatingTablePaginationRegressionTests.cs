using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Pdf;
using Xunit;
using Pig = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests;

public class PdfFloatingTablePaginationRegressionTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MinimumHeightClearanceIncludesOffsetFloatAndDynamicContent(bool dynamicContent) {
        var table = Floating(320, 50);
        table.Position = new PdfTablePosition(verticalOffset: global::OfficeIMO.OfficeLength.Points(80));
        using var pdf = Pig.Open(PdfDocument.Create(Options(240))
            .Table(new[] { new[] { "floating" } }, style: table)
            .Flow(flow => {
                if (dynamicContent) flow.Deferred(_ => nested => nested.AddParagraph(p => p.AddText("first")));
                else flow.AddParagraph(p => p.AddText("first"));
            }, new PdfFlowOptions { MinimumRemainingHeight = global::OfficeIMO.OfficeLength.Points(100) }).ToBytes());
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), w => w.Text == "first");
        Assert.Contains(pdf.GetPage(2).GetWords(), w => w.Text == "first");
    }
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FloatClearanceRechecksFlowMinimumHeight(bool keepTogether) {
        using var pdf = Pig.Open(PdfDocument.Create(Options(240))
            .Table(new[] { new[] { "floating" } }, style: Floating(320, 100))
            .Flow(flow => flow.AddParagraph(p => p.AddText("first")),
                new PdfFlowOptions { KeepTogether = keepTogether, MinimumRemainingHeight = global::OfficeIMO.OfficeLength.Points(100) }).ToBytes());
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), w => w.Text == "first");
        Assert.Contains(pdf.GetPage(2).GetWords(), w => w.Text == "first");
    }
    [Fact]
    public void SkippedFlowDoesNotConsumeFloatClearance() {
        var options = Options(240);
        var table = Floating(120, 100);
        byte[] baseline = PdfDocument.Create(options).Table(new[] { new[] { "floating" } }, style: table)
            .Paragraph(p => p.AddText("following")).ToBytes();
        byte[] actual = PdfDocument.Create(options).Table(new[] { new[] { "floating" } }, style: table)
            .Flow(flow => flow.AddParagraph(p => p.AddText("skipped")).Spacer(global::OfficeIMO.OfficeLength.Points(70)),
                new PdfFlowOptions { OverflowBehavior = PdfFlowOverflowBehavior.Skip })
            .Paragraph(p => p.AddText("following")).ToBytes();
        using var expectedPdf = Pig.Open(baseline);
        using var actualPdf = Pig.Open(actual);
        Assert.Equal(expectedPdf.GetPage(1).GetWords().Single(w => w.Text == "following").BoundingBox,
            actualPdf.GetPage(1).GetWords().Single(w => w.Text == "following").BoundingBox);
        Assert.DoesNotContain(actualPdf.GetPage(1).GetWords(), w => w.Text == "skipped");
    }
    [Theory]
    [InlineData(PdfFlowOverflowBehavior.Skip)]
    [InlineData(PdfFlowOverflowBehavior.StopDocument)]
    public void FloatClearancePrecedesNonContinuingOverflowDecision(PdfFlowOverflowBehavior behavior) {
        using var pdf = Pig.Open(PdfDocument.Create(Options(240))
            .Table(new[] { new[] { "floating" } }, style: Floating(320, 100))
            .Flow(flow => flow.AddParagraph(p => p.AddText("first")).Spacer(global::OfficeIMO.OfficeLength.Points(70)),
                new PdfFlowOptions { OverflowBehavior = behavior })
            .Paragraph(p => p.AddText("following")).ToBytes());
        Assert.DoesNotContain(pdf.GetPages().SelectMany(p => p.GetWords()), w => w.Text == "first");
        Assert.Equal(behavior == PdfFlowOverflowBehavior.Skip,
            pdf.GetPages().SelectMany(p => p.GetWords()).Any(w => w.Text == "following"));
    }
    [Fact]
    public void StaticConstrainedFlowRemeasuresSpacingAfterFloatClearance() {
        using var pdf = Pig.Open(PdfDocument.Create(Options(240))
            .Table(new[] { new[] { "floating" } }, style: Floating(320, 50))
            .Flow(flow => flow.AddParagraph(p => p.AddText("first"), style: new PdfParagraphStyle { LineHeight = 1, SpacingBefore = global::OfficeIMO.OfficeLength.Points(50), SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) })
                .Spacer(global::OfficeIMO.OfficeLength.Points(80)).AddParagraph(p => p.AddText("last"), style: new PdfParagraphStyle { LineHeight = 1, SpacingAfter = global::OfficeIMO.OfficeLength.Points(0) }),
                new PdfFlowOptions { KeepTogether = true }).ToBytes());
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "first");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "first");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "last");
    }

    [Fact]
    public void ReplayedConstrainedFlowClearsLowerFloatWithNewHeight() {
        var lower = Floating(320, 30); lower.Position = new PdfTablePosition(verticalOffset: global::OfficeIMO.OfficeLength.Points(70));
        using var pdf = Pig.Open(PdfDocument.Create(Options(240))
            .Table(new[] { new[] { "top" } }, style: Floating(320, 30))
            .Table(new[] { new[] { "lower" } }, style: lower)
            .Deferred(context => flow => flow.AddParagraph(p => p.AddText(context.IsAtPageTop ? "first" : "first\none\ntwo\nthree\nfour\nfive"))
                .Spacer(global::OfficeIMO.OfficeLength.Points(30)).AddParagraph(p => p.AddText("last")), new PdfFlowOptions { KeepTogether = true }).ToBytes());
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "first");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "first");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "last");
    }

    [Fact]
    public void ConstrainedFlowMovesAsAWholePastFloat() {
        using var pdf = Pig.Open(PdfDocument.Create(Options(240))
            .Table(new[] { new[] { "floating" } }, style: Floating(320, 100))
            .Flow(flow => flow.AddParagraph(p => p.AddText("first")).Spacer(global::OfficeIMO.OfficeLength.Points(50)).AddParagraph(p => p.AddText("last")),
                new PdfFlowOptions { KeepTogether = true }).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "first" || word.Text == "last");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "first");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "last");
    }

    [Fact]
    public void FloatKeepWithNextIncludesAfterSpacing() {
        using var pdf = Pig.Open(PdfDocument.Create(Options(240))
            .Table(new[] { new[] { "floating" } }, style: Floating(320, 90))
            .Paragraph(p => p.AddText("first"), style: new PdfParagraphStyle { KeepWithNext = true, SpacingAfter = global::OfficeIMO.OfficeLength.Points(50) })
            .Paragraph(p => p.AddText("next")).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "first");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "first");
        Assert.Contains(pdf.GetPage(2).GetWords(), word => word.Text == "next");
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void FullPageFloatMovesFirstParagraphLineToNextPage(bool deferred) {
        var style = Floating(320, 150);
        var document = PdfDocument.Create(Options(240));
        if (deferred) document.TableDeferred(() => new[] { new[] { "floating" } }, batchSize: 1, style: style);
        else document.Table(new[] { new[] { "floating" } }, style: style);
        using var pdf = Pig.Open(document.Paragraph(p => p.AddText("following")).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "following");
        Assert.Equal(40, pdf.GetPage(2).GetWords().Single(word => word.Text == "following").BoundingBox.Left, 1);
    }

    [Fact]
    public void FloatingBookmarksMoveTogetherWithTheirFollowingParagraph() {
        var document = PdfDocument.Create(Options(240))
            .Table(new[] { new[] { "floating" } }, style: Floating(320, 150));
        for (int index = 0; index < 200; index++)
            document.Bookmark("destination-" + index);

        byte[] bytes = document.Paragraph(paragraph => paragraph.AddText("following")).ToBytes();
        var destinations = PdfInspector.Inspect(bytes).NamedDestinations
            .Where(destination => destination.Name.StartsWith("destination-"))
            .ToArray();

        Assert.Equal(200, destinations.Length);
        Assert.All(destinations, destination => Assert.Equal(2, destination.PageNumber));
    }


    private static PdfCreateOptions Options(double height = 500) => new() {
        PageWidth = global::OfficeIMO.OfficeLength.Points(400), PageHeight = global::OfficeIMO.OfficeLength.Points(height), MarginLeft = global::OfficeIMO.OfficeLength.Points(40), MarginRight = global::OfficeIMO.OfficeLength.Points(40), MarginTop = global::OfficeIMO.OfficeLength.Points(40), MarginBottom = global::OfficeIMO.OfficeLength.Points(40)
    };
    private static PdfTableStyle Floating(double width = 120, double height = 80) => new() {
        HeaderRowCount = 0, ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(width) }, MinRowHeight = global::OfficeIMO.OfficeLength.Points(height),
        Position = new PdfTablePosition()
    };

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PageMovedFloatRestoresDestinationFlow(bool deferred, bool keepTogether) {
        var style = Floating(); style.KeepTogether = keepTogether;
        var document = PdfDocument.Create(Options()).Paragraph(p => p.AddText("intro")).Spacer(350);
        if (deferred) document.TableDeferred(() => new[] { new[] { "floating" } }, batchSize: 1, style: style);
        else document.Table(new[] { new[] { "floating" } }, style: style);
        using var pdf = Pig.Open(document.Paragraph(p => p.AddText("following")).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        var following = pdf.GetPage(2).GetWords().Single(w => w.Text == "following");
        Assert.True(following.BoundingBox.Left >= 160);
        Assert.True(following.BoundingBox.Top > 430);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LeadingTabIncludesAdvanceWhenAvoidingFloat(bool inline) {
        var style = Floating(200); style.Position = new PdfTablePosition(horizontalAlignment: OfficeAlignment.Right);
        var paragraphStyle = new PdfParagraphStyle { DefaultTabStopWidth = global::OfficeIMO.OfficeLength.Points(100) };
        using var pdf = Pig.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: style)
            .Paragraph(p => {
                p.AddText("\t");
                if (inline) p.AddInline(new PdfInlineBox(global::OfficeIMO.OfficeLength.Points(50), global::OfficeIMO.OfficeLength.Points(20), background: global::OfficeIMO.OfficeColor.Black));
                p.AddText("following");
            }, style: paragraphStyle).ToBytes());
        var following = pdf.GetPage(1).GetWords().Single(w => w.Text == "following");
        Assert.True(following.BoundingBox.Top <= 381, following.BoundingBox.ToString());
    }

    [Fact]
    public void EntirePanelAvoidsOffsetFloat() {
        var style = Floating(); style.Position = new PdfTablePosition(verticalOffset: global::OfficeIMO.OfficeLength.Points(25));
        using var pdf = Pig.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: style)
            .Panel(panel => panel.AddParagraph(p => p.AddText("one\ntwo\nthree\nfour\nfive")),
                new PdfPanelStyle { Background = global::OfficeIMO.OfficeColor.Black, KeepTogether = false }).ToBytes());
        Assert.True(pdf.GetPage(1).GetWords().Single(w => w.Text == "one").BoundingBox.Top <= 356);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SectionDestinationFollowsFinalHeadingPlacement(bool newPage) {
        var document = PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: Floating());
        if (newPage) document.Spacer(405);
        byte[] bytes = document.Section("Heading", _ => { }, new PdfSectionOptions { DestinationName = "target" }).ToBytes();
        using var pdf = Pig.Open(bytes);
        var destination = Assert.Single(PdfInspector.Inspect(bytes).NamedDestinations);
        Assert.Equal(newPage ? 2 : 1, destination.PageNumber);
        var heading = pdf.GetPage(destination.PageNumber!.Value).GetWords().Single(w => w.Text == "Heading");
        Assert.InRange(destination.DestinationTop!.Value - heading.BoundingBox.Top, 0, 12);
        if (!newPage) Assert.True(destination.DestinationTop <= 380);
    }

    [Fact]
    public void WidowTransferredLinesDiscardPriorPageFloatOffsets() {
        using var pdf = Pig.Open(PdfDocument.Create(Options(240)).Table(new[] { new[] { "floating" } }, style: Floating(200, 145))
            .Paragraph(p => p.AddText(string.Join("\n", Enumerable.Range(0, 12).Select(i => "word" + i))),
                style: new PdfParagraphStyle { WidowControl = true }).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        var words = pdf.GetPage(2).GetWords().Where(w => w.Text.StartsWith("word")).ToList();
        Assert.True(words.Count >= 2);
        Assert.All(words, word => Assert.Equal(40, word.BoundingBox.Left, 1));
    }

    [Fact]
    public void WidowTransferredTabsRetainParagraphStopColumn() {
        var style = new PdfParagraphStyle { WidowControl = true };
        style.TabStops.Add(new PdfTabStop(global::OfficeIMO.OfficeLength.Points(240)));
        using var pdf = Pig.Open(PdfDocument.Create(Options(240)).Table(new[] { new[] { "floating" } }, style: Floating(200, 145))
            .Paragraph(p => p.AddText(string.Join("\n", Enumerable.Range(0, 12).Select(i => "\tword" + i))), style: style).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        var words = pdf.GetPage(2).GetWords().Where(w => w.Text.StartsWith("word")).ToList();
        Assert.True(words.Count >= 2);
        Assert.All(words, word => Assert.Equal(280, word.BoundingBox.Left, 1));
    }
}
