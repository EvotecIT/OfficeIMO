using System.Collections.Concurrent;
using System.Threading.Tasks;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDocumentCapabilityLifetimeTests {
    [Fact]
    public void LoadedDocumentCapabilitiesKeepStableIdentity() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Capability lifetime"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(bytes);

        Assert.Same(document.Pages, document.Pages);
        Assert.Same(document.Render, document.Render);
        Assert.Same(document.Resources, document.Resources);
        Assert.Same(document.Text, document.Text);
        Assert.Same(document.Images, document.Images);
        Assert.Same(document.Stamp, document.Stamp);
        Assert.Same(document.Forms, document.Forms);
        Assert.Same(document.Attachments, document.Attachments);
        Assert.Same(document.Bookmarks, document.Bookmarks);
        Assert.Same(document.Annotations, document.Annotations);
        Assert.Same(document.JavaScript, document.JavaScript);
        Assert.Same(document.Security, document.Security);
        Assert.Same(document.Redactions, document.Redactions);
        Assert.Same(document.Optimization, document.Optimization);
        Assert.Same(document.Proof, document.Proof);
    }

    [Fact]
    public void LoadedDocumentAuthoringDefaultsRemainIsolated() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Source-backed option isolation"))
            .ToBytes();
        PdfDocument first = PdfDocument.Load(bytes);
        PdfDocument second = PdfDocument.Load(bytes);

        Assert.NotSame(first.Options, second.Options);

        first.Options.MarginLeft = 11;

        Assert.Equal(11, first.Options.MarginLeft);
        Assert.Equal(72, second.Options.MarginLeft);
    }

    [Fact]
    public void ConcurrentFirstAccessPublishesOnePagesCapability() {
        byte[] bytes = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("Concurrent capability lifetime"))
            .ToBytes();
        PdfDocument document = PdfDocument.Load(bytes);
        var capabilities = new ConcurrentBag<PdfDocumentPages>();

        Parallel.For(0, 64, _ => capabilities.Add(document.Pages));

        PdfDocumentPages expected = Assert.Single(capabilities.Distinct());
        Assert.Same(expected, document.Pages);
    }
}
