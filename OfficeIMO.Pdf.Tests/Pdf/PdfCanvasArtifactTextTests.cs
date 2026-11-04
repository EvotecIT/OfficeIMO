using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCanvasArtifactTextTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ArtifactTextDoesNotRegisterLinkPaintWhileOrdinaryLinksRemainTagged(bool textBox) {
        const string uri = "https://example.test/canvas-link";
        byte[] Render(bool artifact) => PdfDocument.Create(new PdfCreateOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            CompressContentStreams = false
        }).Canvas(canvas => {
            void Paint(PdfPageCanvas target) {
                var runs = new[] { PdfTextRun.Link("CanvasLinkMarker", uri) };
                if (textBox) target.TextBox(runs, global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(50D), global::OfficeIMO.OfficeLength.Points(180D), global::OfficeIMO.OfficeLength.Points(40D));
                else target.Text(runs, global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(50D), global::OfficeIMO.OfficeLength.Points(180D), global::OfficeIMO.OfficeLength.Points(40D));
            }
            if (artifact) canvas.Artifact(Paint);
            else Paint(canvas);
        }).ToBytes();

        byte[] decorative = Render(artifact: true);
        string artifactSyntax = Encoding.ASCII.GetString(decorative);
        Assert.Contains("/Artifact << /Type /Layout >> BDC", artifactSyntax);
        Assert.DoesNotContain("/Artifact BMC", artifactSyntax);
        Assert.DoesNotContain("/MCID", artifactSyntax);
        Assert.Empty(PdfInspector.Inspect(decorative).LinkUris);

        byte[] ordinary = Render(artifact: false);
        Assert.Contains("/Link << /MCID", Encoding.ASCII.GetString(ordinary));
        Assert.Contains(uri, PdfInspector.Inspect(ordinary).LinkUris);
        Assert.Contains("CanvasLinkMarker", PdfReadDocument.Open(ordinary).ExtractText());
    }

    [Fact]
    public void SecondaryTextPaintIsALayoutArtifactWithoutDuplicateExtraction() {
        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            CompressContentStreams = false
        }).Canvas(canvas => canvas
            .SuppressTextExtraction(paint => paint.Text("DuplicatePaint", global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(50D), global::OfficeIMO.OfficeLength.Points(180D), global::OfficeIMO.OfficeLength.Points(40D)))
            .Text("Meaningful", global::OfficeIMO.OfficeLength.Points(20D), global::OfficeIMO.OfficeLength.Points(100D), global::OfficeIMO.OfficeLength.Points(180D), global::OfficeIMO.OfficeLength.Points(40D))).ToBytes();

        string syntax = Encoding.ASCII.GetString(pdf);
        Assert.Contains("/Artifact << /Type /Layout >> BDC\n/Span << /ActualText", syntax);
        Assert.DoesNotContain("/Artifact BMC", syntax);
        string text = PdfReadDocument.Open(pdf).ExtractText();
        Assert.DoesNotContain("DuplicatePaint", text);
        Assert.Contains("Meaningful", text);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SourceStructuredDecorativePaintIsALayoutArtifact(bool explicitlyDecorative) {
        var drawing = new OfficeDrawing(180D, 40D).AddText("Decoration", 0D, 0D, 180D, 40D);
        drawing.Elements[0].SourceElementIds = new[] { "decoration" };
        var structure = new Dictionary<string, PdfCanvasSourceContent>();
        if (explicitlyDecorative) {
            structure.Add("decoration", new PdfCanvasSourceContent(Array.Empty<PdfCanvasStructureStep>(), 0, artifact: true));
        }
        byte[] pdf = PdfDocument.Create(new PdfCreateOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            CompressContentStreams = false
        }).Canvas(canvas => canvas.SourceStructuredDrawing(drawing, 180D, 40D, structure)).ToBytes();

        string syntax = Encoding.ASCII.GetString(pdf);
        Assert.Contains("/Artifact << /Type /Layout >> BDC", syntax);
        Assert.DoesNotContain("/Artifact BMC", syntax);
        Assert.DoesNotContain("/MCID", syntax);
        Assert.DoesNotContain("Decoration", PdfReadDocument.Open(pdf).ExtractText());
    }
}
