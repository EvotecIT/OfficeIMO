using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ParsedInspectionIgnoresFeatureNamesInsideOpaquePayloads(bool inStream) {
        const string markers = "/FT /AcroForm /ByteRange /Sig /Annots /Outlines /PageMode /PageLabels /Names /Dests /OpenAction /ViewerPreferences /StructElem /Metadata /OutputIntent /EmbeddedFile";
        byte[] pdf = BuildOpaqueMarkerPdf(markers, inStream);
        PdfDocumentInfo info = PdfInspector.Inspect(pdf);
        Assert.False(info.HasForms);
        Assert.False(info.HasSignatures);
        Assert.False(info.Security.HasSignatures);
        Assert.False(info.Security.HasByteRange);
        Assert.False(PdfSyntax.ReadDocumentSecurityInfo(pdf).HasSignatures);
        Assert.False(PdfSyntax.HasSignatureMarkers(pdf));
        Assert.False(PdfSyntax.HasFormMarkers(pdf));
        Assert.False(PdfSyntax.HasOutputIntentMarkers(pdf));
        Assert.False(PdfSyntax.HasXmpMetadataMarkers(pdf));
        Assert.False(PdfSyntax.HasPageLabelMarkers(pdf));
        Assert.False(info.HasAnnotations);
        Assert.False(info.HasOutlines);
        Assert.False(info.HasCatalogViewSettings);
        Assert.False(info.HasPageLabels);
        Assert.False(info.HasCatalogNameTrees);
        Assert.False(info.HasNamedDestinations);
        Assert.False(info.HasOpenActions);
        Assert.False(info.HasViewerPreferences);
        Assert.False(info.HasTaggedContent);
        Assert.False(info.HasXmpMetadata);
        Assert.False(info.HasOutputIntents);
        Assert.False(info.HasEmbeddedFiles);
    }

    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    public void SecurityRewriteDoesNotTreatPayloadTextAsAForm(PdfStandardEncryptionAlgorithm algorithm) {
        const string title = "/FT /ByteRange /Sig";
        byte[] pdf = BuildOpaqueMarkerPdf(title, inStream: false);
        var encryption = new PdfStandardEncryptionOptions("open") { OwnerPassword = "owner", Algorithm = algorithm };
        var protectedPdf = PdfDocument.Load(pdf).Security.Encrypt(encryption);
        var unlockedPdf = protectedPdf.ToDocument().Security.Decrypt("owner");
        Assert.False(PdfInspector.Inspect(protectedPdf.Pdf, new PdfLoadOptions { Password = "owner" }).HasForms);
        Assert.False(PdfSyntax.HasSignatureMarkers(protectedPdf.Pdf, new PdfLoadOptions { Password = "owner" }));
        Assert.False(PdfInspector.Inspect(unlockedPdf.Pdf).HasForms);
        Assert.Equal(title, PdfInspector.Inspect(unlockedPdf.Pdf).Metadata.Title);
    }

    [Fact]
    public void UnparseableMarkerProbeRemainsConservative() {
        byte[] incomplete = Encoding.ASCII.GetBytes("%PDF-1.7\n/FT /Sig /OutputIntents\n%%EOF");
        Assert.True(PdfSyntax.HasFormMarkers(incomplete));
        Assert.True(PdfSyntax.HasSignatureMarkers(incomplete));
        Assert.True(PdfSyntax.HasOutputIntentMarkers(incomplete));
    }

    private static byte[] BuildOpaqueMarkerPdf(string markers, bool inStream) {
        string content = inStream ? "q\n% " + markers + "\nQ" : "q\nQ";
        string title = inStream ? "Opaque payload" : markers;
        return Encoding.ASCII.GetBytes(string.Join("\n", new[] {
            "%PDF-1.7", "1 0 obj", "<< /Type /Catalog /Pages 2 0 R >>", "endobj",
            "2 0 obj", "<< /Type /Pages /Count 1 /Kids [3 0 R] >>", "endobj",
            "3 0 obj", "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 200 200] /Contents 4 0 R >>", "endobj",
            "4 0 obj", "<< /Length " + content.Length + " >>", "stream", content, "endstream", "endobj",
            "5 0 obj", "<< /Title (" + title + ") >>", "endobj",
            "trailer", "<< /Root 1 0 R /Info 5 0 R /Size 6 >>", "%%EOF"
        }));
    }
}
