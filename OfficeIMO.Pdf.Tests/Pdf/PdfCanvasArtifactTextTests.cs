using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfCanvasArtifactTextTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ArtifactTextDoesNotRegisterLinkPaintWhileOrdinaryLinksRemainTagged(bool textBox) {
        const string uri = "https://example.test/canvas-link";
        byte[] Render(bool artifact) => PdfDocument.Create(new PdfOptions {
            TaggedStructureMode = PdfTaggedStructureMode.CatalogMarkers,
            CompressContentStreams = false
        }).Canvas(canvas => {
            void Paint(PdfPageCanvas target) {
                var runs = new[] { PdfTextRun.Link("CanvasLinkMarker", uri) };
                if (textBox) target.TextBox(runs, 20D, 50D, 180D, 40D);
                else target.Text(runs, 20D, 50D, 180D, 40D);
            }
            if (artifact) canvas.Artifact(Paint);
            else Paint(canvas);
        }).ToBytes();

        byte[] decorative = Render(artifact: true);
        string artifactSyntax = Encoding.ASCII.GetString(decorative);
        Assert.Contains("/Artifact BMC", artifactSyntax);
        Assert.DoesNotContain("/MCID", artifactSyntax);
        Assert.Empty(PdfInspector.Inspect(decorative).LinkUris);

        byte[] ordinary = Render(artifact: false);
        Assert.Contains("/Link << /MCID", Encoding.ASCII.GetString(ordinary));
        Assert.Contains(uri, PdfInspector.Inspect(ordinary).LinkUris);
        Assert.Contains("CanvasLinkMarker", PdfReadDocument.Open(ordinary).ExtractText());
    }
}
