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
    public void FullPageFloatMovesFirstParagraphLineToNextPage(bool deferred) {
        var style = Floating(320, 150);
        var document = PdfDocument.Create(Options(240));
        if (deferred) document.TableDeferred(() => new[] { new[] { "floating" } }, batchSize: 1, style: style);
        else document.Table(new[] { new[] { "floating" } }, style: style);
        using var pdf = Pig.Open(document.Paragraph(p => p.Text("following")).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.DoesNotContain(pdf.GetPage(1).GetWords(), word => word.Text == "following");
        Assert.Equal(40, pdf.GetPage(2).GetWords().Single(word => word.Text == "following").BoundingBox.Left, 1);
    }

    private static PdfOptions Options(double height = 500) => new() {
        PageWidth = 400, PageHeight = height, MarginLeft = 40, MarginRight = 40, MarginTop = 40, MarginBottom = 40
    };
    private static PdfTableStyle Floating(double width = 120, double height = 80) => new() {
        HeaderRowCount = 0, ColumnWidthPoints = new List<double?> { width }, MinRowHeight = height,
        Position = new PdfTablePosition()
    };

    [Theory]
    [InlineData(false, false)]
    [InlineData(false, true)]
    [InlineData(true, false)]
    [InlineData(true, true)]
    public void PageMovedFloatRestoresDestinationFlow(bool deferred, bool keepTogether) {
        var style = Floating(); style.KeepTogether = keepTogether;
        var document = PdfDocument.Create(Options()).Paragraph(p => p.Text("intro")).Spacer(350);
        if (deferred) document.TableDeferred(() => new[] { new[] { "floating" } }, batchSize: 1, style: style);
        else document.Table(new[] { new[] { "floating" } }, style: style);
        using var pdf = Pig.Open(document.Paragraph(p => p.Text("following")).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        var following = pdf.GetPage(2).GetWords().Single(w => w.Text == "following");
        Assert.True(following.BoundingBox.Left >= 160);
        Assert.True(following.BoundingBox.Top > 430);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void LeadingTabIncludesAdvanceWhenAvoidingFloat(bool inline) {
        var style = Floating(200); style.Position = new PdfTablePosition(horizontalAlignment: PdfAlign.Right);
        var paragraphStyle = new PdfParagraphStyle { DefaultTabStopWidth = 100 };
        using var pdf = Pig.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: style)
            .Paragraph(p => {
                p.Text("\t");
                if (inline) p.Inline(new PdfInlineBox(50, 20, background: PdfColor.Black));
                p.Text("following");
            }, style: paragraphStyle).ToBytes());
        var following = pdf.GetPage(1).GetWords().Single(w => w.Text == "following");
        Assert.True(following.BoundingBox.Top <= 381, following.BoundingBox.ToString());
    }

    [Fact]
    public void EntirePanelAvoidsOffsetFloat() {
        var style = Floating(); style.Position = new PdfTablePosition(verticalOffset: 25);
        using var pdf = Pig.Open(PdfDocument.Create(Options()).Table(new[] { new[] { "floating" } }, style: style)
            .Panel(panel => panel.Paragraph(p => p.Text("one\ntwo\nthree\nfour\nfive")),
                new PdfPanelStyle { Background = PdfColor.Black, KeepTogether = false }).ToBytes());
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
            .Paragraph(p => p.Text(string.Join("\n", Enumerable.Range(0, 12).Select(i => "word" + i))),
                style: new PdfParagraphStyle { WidowControl = true }).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        var words = pdf.GetPage(2).GetWords().Where(w => w.Text.StartsWith("word")).ToList();
        Assert.True(words.Count >= 2);
        Assert.All(words, word => Assert.Equal(40, word.BoundingBox.Left, 1));
    }

    [Fact]
    public void WidowTransferredTabsRetainParagraphStopColumn() {
        var style = new PdfParagraphStyle { WidowControl = true };
        style.TabStops.Add(new PdfTabStop(240));
        using var pdf = Pig.Open(PdfDocument.Create(Options(240)).Table(new[] { new[] { "floating" } }, style: Floating(200, 145))
            .Paragraph(p => p.Text(string.Join("\n", Enumerable.Range(0, 12).Select(i => "\tword" + i))), style: style).ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        var words = pdf.GetPage(2).GetWords().Where(w => w.Text.StartsWith("word")).ToList();
        Assert.True(words.Count >= 2);
        Assert.All(words, word => Assert.Equal(280, word.BoundingBox.Left, 1));
    }
}
