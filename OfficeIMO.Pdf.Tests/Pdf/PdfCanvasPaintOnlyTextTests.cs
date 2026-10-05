using System;
using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCanvasPaintOnlyTextTests {
    [Fact]
    public void PaintOnlyScopeSuppressesNestedLogicalTextWithoutChangingOrdinaryText() {
        byte[] pdf = PdfDocument.Create(new PdfOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            CompressContentStreams = false
        }).Canvas(canvas => {
            canvas.Text("VisibleMarker", 20D, 20D, 180D, 20D);
            canvas.SuppressTextExtraction(paint => paint.ActualText("HiddenLogicalMarker", child =>
                child.Text("PaintedMarker", 20D, 50D, 180D, 20D)));
            canvas.Text("FollowingMarker", 20D, 80D, 180D, 20D);
        }).ToBytes();

        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.Contains("VisibleMarker", text);
        Assert.Contains("FollowingMarker", text);
        Assert.DoesNotContain("HiddenLogicalMarker", text);
        Assert.DoesNotContain("PaintedMarker", text);
        string syntax = Encoding.ASCII.GetString(pdf);
        Assert.Contains("/StructTreeRoot", syntax);
        Assert.Contains("/Artifact BMC\n/Span << /ActualText <> >> BDC\n", syntax);
        Assert.Throws<ArgumentException>(() => new PdfPageCanvas().ActualText(string.Empty,
            child => child.Text("Marker", 20D, 20D, 180D, 20D)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PaintOnlyLinkedTextRetainsItsAnnotationWithoutTaggedPaintInsideArtifact(bool textBox) {
        const string uri = "https://example.test/secondary";
        byte[] pdf = PdfDocument.Create(new PdfOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            CompressContentStreams = false
        }).Canvas(canvas => canvas.SuppressTextExtraction(paint => {
            var runs = new[] { PdfTextRun.Link("SecondaryMarker", uri) };
            if (textBox) paint.TextBox(runs, 20D, 50D, 180D, 40D);
            else paint.Text(runs, 20D, 50D, 180D, 40D);
        })).ToBytes();

        string syntax = Encoding.ASCII.GetString(pdf);
        Assert.Contains("/Artifact BMC\n/Span << /ActualText <> >> BDC\n", syntax);
        Assert.DoesNotContain("/MCID", syntax);
        Assert.Contains("/OBJR", syntax);
        Assert.Contains("/StructParent ", syntax);
        Assert.Contains(uri, PdfInspector.Inspect(pdf).LinkUris);
        Assert.DoesNotContain("SecondaryMarker", PdfReadDocument.Open(pdf).ExtractText());
    }
}
