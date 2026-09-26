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
}
