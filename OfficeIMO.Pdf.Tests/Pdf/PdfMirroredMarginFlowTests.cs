using System;
using System.Linq;
using OfficeIMO.Pdf;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfMirroredMarginFlowTests {
    [Theory]
    [InlineData("paragraph")]
    [InlineData("table")]
    [InlineData("list")]
    [InlineData("row")]
    [InlineData("container")]
    public void MirroredMarginsFollowAutomaticFlowContinuation(string mode) {
        string[] markers = Enumerable.Range(0, 30).Select(index => "M" + index.ToString("D2")).ToArray();
        var document = PdfDocument.Create(builder => builder.Content(content => {
            switch (mode) {
                case "paragraph":
                    content.Paragraph(paragraph => paragraph.Text(string.Join("\n", markers)));
                    break;
                case "table":
                    content.Table(markers.Select(marker => new[] { marker }), style: new PdfTableStyle {
                        HeaderRowCount = 0, CellPaddingX = 0, CellPaddingY = 2
                    });
                    break;
                case "list":
                    content.Bullets(new[] { string.Join("\n", markers) });
                    break;
                case "row":
                    content.Row(row => row
                        .PercentColumn(50, column => column.Paragraph(paragraph => paragraph.Text(string.Join("\n", markers))))
                        .PercentColumn(50, column => column.Text("RightColumn")));
                    break;
                case "container":
                    content.Element(outer => outer.Padding(2, 7).Content(nested => nested.Element(inner => inner.Padding(2, 5).Content(body => {
                        foreach (string marker in markers) body.Text(marker);
                    }))));
                    content.Text("AfterContainer");
                    break;
            }
        }), MirroredOptions());

        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.True(pdf.NumberOfPages >= 2);
        double offset = FindMarkerX(pdf.GetPage(1), "M00")!.Value - 40;
        int seen = 0;
        for (int pageNumber = 1; pageNumber <= pdf.NumberOfPages; pageNumber++) {
            var page = pdf.GetPage(pageNumber);
            double margin = pageNumber % 2 == 0 ? 70 : 40;
            foreach (string marker in markers) {
                double? markerX = FindMarkerX(page, marker);
                if (markerX.HasValue) {
                    Assert.Equal(margin + offset, markerX.Value, 3);
                    seen++;
                }
            }
            double? afterX = FindMarkerX(page, "AfterContainer");
            if (afterX.HasValue) Assert.Equal(margin, afterX.Value, 3);
        }
        Assert.Equal(markers.Length, seen);
    }

    [Fact]
    public void MirroredMarginCloneRetainsParityWithoutChangingItsSourceMargins() {
        var options = MirroredOptions();
        options.PageNumberStart = 2;
        var document = PdfDocument.Create(builder => builder.Content(content =>
            content.Text("EvenFirst").PageBreak().Text("OddSecond")), options.Clone());
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(70, pdf.GetPage(1).GetWords().Single().BoundingBox.Left, 3);
        Assert.Equal(40, pdf.GetPage(2).GetWords().Single().BoundingBox.Left, 3);
        Assert.Equal(40, options.MarginLeft);
        Assert.Equal(70, options.MarginRight);
    }

    [Fact]
    public void MirroredMarginsFollowNumberingAcrossSectionsAndRestarts() {
        var document = PdfDocument.Create(builder => builder
            .Section(page => page.MirrorMargins().Content(content => content.Text("A").PageBreak().Text("B")))
            .Section(page => page.MirrorMargins().Content(content => content.Text("C")))
            .Section(page => page.MirrorMargins().PageNumberStart(2).Content(content => content.Text("D").PageBreak().Text("E")))
            .Content(content => content.Text("Tail")), MirroredOptions());
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(6, pdf.NumberOfPages);
        double[] expected = { 40, 70, 40, 70, 40, 70 };
        string[] markers = { "A", "B", "C", "D", "E", "Tail" };
        for (int index = 0; index < markers.Length; index++)
            Assert.Equal(expected[index], FindMarkerX(pdf.GetPage(index + 1), markers[index])!.Value, 3);
    }

    [Fact]
    public void MirroredMarginsUseVisibleParityAfterASectionPaddingPage() {
        var options = MirroredOptions();
        options.PageStartParity = PdfPageParity.Odd;
        var document = PdfDocument.Create(builder => builder
            .Section(page => page.Content(content => content.Text("First")))
            .Section(page => page.PageNumberStart(2).Content(content => content.Text("EvenNumber"))), options);
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(3, pdf.NumberOfPages);
        Assert.Empty(pdf.GetPage(2).Letters);
        Assert.Equal(70, FindMarkerX(pdf.GetPage(3), "EvenNumber")!.Value, 3);
    }

    [Fact]
    public void MirroredMarginsPreserveAbsoluteCanvasCoordinates() {
        var document = PdfDocument.Create(builder => builder.Content(content => content
            .Text("First").PageBreak().Text("Flow")
            .Canvas(canvas => canvas.Text("Canvas", 15, 60, 100, 20))), MirroredOptions());
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(70, FindMarkerX(pdf.GetPage(2), "Flow")!.Value, 3);
        Assert.Equal(15, FindMarkerX(pdf.GetPage(2), "Canvas")!.Value, 3);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void MirroredMarginsFollowKeepTogetherMoves(bool table) {
        var document = PdfDocument.Create(builder => builder.Content(content => {
            content.Text("Seed").Spacer(130);
            if (table) content.Table(new[] { new[] { "Moved" }, new[] { "Together" } },
                style: new PdfTableStyle { HeaderRowCount = 0, CellPaddingX = 0, KeepTogether = true });
            else content.Text("Moved\nTogether", style: new PdfParagraphStyle { KeepTogether = true });
        }), MirroredOptions());
        using var pdf = PdfPigDocument.Open(document.ToBytes());
        Assert.Equal(2, pdf.NumberOfPages);
        Assert.Equal(70, FindMarkerX(pdf.GetPage(2), "Moved")!.Value, 3);
    }

    [Theory]
    [InlineData("text")]
    [InlineData("checkbox")]
    [InlineData("choice")]
    [InlineData("radio")]
    public void MirroredMarginsMoveFlowFormWidgetRectanglesWithTheirPage(string kind) {
        var document = PdfDocument.Create(builder => builder.Content(content => {
            content.Text("Seed").Spacer(130);
            switch (kind) {
                case "text": content.TextField("Field", width: 80); break;
                case "checkbox": content.CheckBox("Field"); break;
                case "choice": content.ChoiceField("Field", new[] { "A" }, width: 80); break;
                case "radio": content.RadioButtonGroup("Field", new[] { "A" }); break;
            }
        }), MirroredOptions());
        var widget = Assert.Single(PdfInspector.Inspect(document.ToBytes()).FormFieldsByName["Field"].Widgets);
        Assert.Equal(2, widget.PageNumber);
        Assert.Equal(70, widget.X1, 3);
    }

    private static PdfOptions MirroredOptions() => new() {
        PageWidth = 300, PageHeight = 200,
        MarginLeft = 40, MarginRight = 70, MarginTop = 20, MarginBottom = 20,
        MirrorMargins = true, DefaultFontSize = 10, CompressContentStreams = false
    };

    private static double? FindMarkerX(UglyToad.PdfPig.Content.Page page, string marker) {
        for (int index = 0; index <= page.Letters.Count - marker.Length; index++) {
            if (string.Concat(page.Letters.Skip(index).Take(marker.Length).Select(letter => letter.Value)) == marker)
                return page.Letters[index].StartBaseLine.X;
        }
        return null;
    }
}
