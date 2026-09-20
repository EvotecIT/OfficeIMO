using System.Collections.Concurrent;
using System.Text;
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

    [Fact]
    public void SplitReadbackRetainsCanonicalSecurityAndRevisionEvidence() {
        PdfDocument source = PdfDocument.Create()
            .Paragraph(paragraph => paragraph.Text("First"))
            .PageBreak()
            .Paragraph(paragraph => paragraph.Text("Second"));

        PdfDocument part = source.Pages.Split()[0];
        PdfReadDocument cached = part.GetReadDocument();
        PdfReadDocument publicReadback = PdfReadDocument.Open(part.ToBytes());

        Assert.Same(cached, part.GetReadDocument());
        Assert.Single(cached.Pages);
        Assert.Equal(publicReadback.Security.HasEncryption, cached.Security.HasEncryption);
        Assert.Equal(publicReadback.Security.HasSignatures, cached.Security.HasSignatures);
        Assert.Equal(publicReadback.Security.HasByteRange, cached.Security.HasByteRange);
        Assert.Equal(publicReadback.Security.RootObjectNumber, cached.Security.RootObjectNumber);
        Assert.Equal(publicReadback.Security.RootObjectGeneration, cached.Security.RootObjectGeneration);
        Assert.Equal(publicReadback.Security.InfoObjectNumber, cached.Security.InfoObjectNumber);
        Assert.Equal(publicReadback.Security.InfoObjectGeneration, cached.Security.InfoObjectGeneration);
        Assert.Equal(publicReadback.Security.HasTrailerId, cached.Security.HasTrailerId);
        Assert.Equal(publicReadback.Security.StartXrefCount, cached.Security.StartXrefCount);
        Assert.Equal(publicReadback.Security.LastStartXrefOffset, cached.Security.LastStartXrefOffset);
        Assert.Equal(publicReadback.Security.StartXrefOffsets, cached.Security.StartXrefOffsets);
        Assert.Equal(publicReadback.Security.RevisionCount, cached.Security.RevisionCount);
        Assert.Equal(publicReadback.Security.HasPreviousRevision, cached.Security.HasPreviousRevision);
        Assert.Equal(publicReadback.Security.HasXrefStreams, cached.Security.HasXrefStreams);
        Assert.Equal(publicReadback.Security.HasObjectStreams, cached.Security.HasObjectStreams);
    }

    [Fact]
    public void SplitReadbackIgnoresStartXrefTextInsideStreamPayload() {
        byte[] source = BuildPdfWithStartXrefStreamPayload();

        PdfDocument part = Assert.Single(PdfDocument.Load(source).Pages.Split());
        PdfReadDocument cached = part.GetReadDocument();

        Assert.Single(cached.Pages);
        Assert.Equal(1, cached.Security.StartXrefCount);
        Assert.True(cached.Security.HasTrailerId);
        Assert.Contains("startxref\n123", Encoding.ASCII.GetString(part.ToBytes()), StringComparison.Ordinal);
    }

    private static byte[] BuildPdfWithStartXrefStreamPayload() {
        byte[] payload = Encoding.ASCII.GetBytes("startxref\n123\n");
        byte[] streamBody = PdfObjectBytes.WrapStreamBody(
            "<< /Length " + payload.Length + " >>",
            payload);
        var objects = new List<byte[]> {
            PdfObjectBytes.WrapIndirectObject(1, "<< /Type /Catalog /Pages 2 0 R >>\n"),
            PdfObjectBytes.WrapIndirectObject(2, "<< /Type /Pages /Kids [3 0 R] /Count 1 >>\n"),
            PdfObjectBytes.WrapIndirectObject(3, "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>\n"),
            PdfObjectBytes.WrapIndirectObject(4, streamBody)
        };
        return PdfFileAssembler.Assemble(objects, 1, 0, PdfFileVersion.Pdf14);
    }
}
