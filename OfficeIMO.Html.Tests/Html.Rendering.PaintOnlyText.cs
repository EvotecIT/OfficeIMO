using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void SecondaryPaintFragmentDoesNotRepeatSearchableTextOrLoseItsLink() {
        const string marker = "BoundaryMarker";
        const string uri = "https://example.test/boundary";
        var font = new OfficeFontInfo("Arial", 16D);
        var visual = new HtmlRenderText(marker, 20D, 384D, 120D, 20D, font, OfficeColor.Black,
            OfficeTextAlignment.Left, 20D, 0, linkUri: uri);
        var projector = new HtmlRenderVisualProjector(16, CancellationToken.None);
        var rendered = new HtmlRenderDocument(HtmlRenderMode.Paged,
            new[] {
                new HtmlRenderPage(1, 400D, 400D, projector.Project(new[] { visual }, 0D, 0D, 400D, 400D, 0D, 0D)),
                new HtmlRenderPage(2, 400D, 400D, projector.Project(new[] { visual }, 0D, 400D, 400D, 400D, 0D, 0D))
            }, new HtmlDiagnosticReport());
        var options = new HtmlToPdfOptions {
            ResourcePolicy = PdfCore.PdfResourcePolicy.CreatePortableDeterministic()
        };

        byte[] pdf = HtmlPdfRenderedConverter.CreatePdf(rendered, options, CancellationToken.None).Document.ToBytes();
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Equal(1, text.Split(new[] { marker }, StringSplitOptions.None).Length - 1);
        Assert.Equal(marker, rendered.Text.Trim());
        var links = PdfCore.PdfInspector.Inspect(pdf).LinkAnnotations.Where(link => link.Uri == uri).ToList();
        Assert.Equal(2, links.Count);
        Assert.All(links, link => {
            Assert.InRange(link.Y1, 0D, 300D);
            Assert.InRange(link.Y2, 0D, 300D);
            Assert.True(link.Y2 > link.Y1);
        });
    }
}
