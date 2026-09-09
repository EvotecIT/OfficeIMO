using System.Text;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfInspectorTests {
    [Theory]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes128)]
    [InlineData(PdfStandardEncryptionAlgorithm.Aes256)]
    [InlineData(PdfStandardEncryptionAlgorithm.LegacyRc4)]
    public void GeneratedEncryptedFormRetainsCompleteSyntaxCoverage(PdfStandardEncryptionAlgorithm algorithm) {
        byte[] pdf = PdfDocument.Create(new PdfOptions().SetEncryption(new("open") { OwnerPassword = "owner", Algorithm = algorithm }))
            .TextField("Name", width: 180, height: 24, value: "Ada").ToBytes();
        PdfSyntax.ParseObjects(pdf, new PdfLoadOptions { Password = "owner" }, out var repair);
        Assert.False(repair.HasUnreadableObjects, string.Join("; ", repair.Diagnostics.Select(item => item.ObjectNumber + ": " + item.Message)));
        using var independent = UglyToad.PdfPig.PdfDocument.Open(pdf, new UglyToad.PdfPig.ParsingOptions { Password = "owner" });
        Assert.Single(independent.GetPages());
    }
    [Theory]
    [InlineData("<< /Type /Sig /ByteRange [0 10 20 30]")]
    [InlineData("<< /FT /Sig /V")]
    [InlineData("<< /ByteRange >>")]
    [InlineData("123 /Sig")]
    [InlineData("<< /Nested << /ByteRange >> >>")]
    [InlineData("<< /Nested [ /ByteRange >>")]
    [InlineData("<< /Custom << /Type /Sig /ByteRange [0 10 20 30] >> /Custom null >>")]
    [InlineData("<< /Nested << /Custom << /Type /Sig >> /Cus#74om null >> >>")]
    public void IncompleteObjectParsingDoesNotErasePotentialSignatures(string damagedObject) {
        string source = Encoding.ASCII.GetString(BuildOpaqueMarkerPdf("Safe metadata", inStream: false));
        byte[] pdf = Encoding.ASCII.GetBytes(source.Replace("trailer", "6 0 obj\n" + damagedObject + "\nendobj\ntrailer"));
        Assert.True(PdfSyntax.HasSignatureMarkers(pdf));
        Assert.True(PdfSyntax.ReadDocumentSecurityInfo(pdf).HasSignatures);
        Assert.True(PdfReadDocument.Open(pdf).Security.HasSignatures);
        Assert.True(PdfInspector.Inspect(pdf).HasSignatures);
        PdfDocumentPreflight preflight = PdfInspector.Preflight(pdf);
        Assert.Contains(preflight.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.IncompleteObjectGraph);
        foreach (PdfMutationOperation operation in System.Enum.GetValues(typeof(PdfMutationOperation))) {
            var plan = PdfMutationPlanner.Plan(preflight, operation);
            Assert.False(plan.CanExecute);
            Assert.Contains("Source.IncompleteObjectGraph", plan.BlockerCodes);
        }
        Assert.ThrowsAny<System.Exception>(() => PdfDocument.Load(pdf).Security.Encrypt(new("open") { OwnerPassword = "owner" }));
    }

    [Theory]
    [InlineData("7 0 << /Type /Sig /ByteRange", 1, 4)]
    [InlineData("7 0 << /Nested << /ByteRange >> >>", 1, 4)]
    [InlineData("7 0 << /Type /Sig >>", 2, 4)]
    [InlineData("7 -1 << /Type /Sig >>", 1, 5)]
    [InlineData("7 0 << /Custom << /Type /Sig >> /Custom null >>", 1, 4)]
    public void IncompleteCompressedObjectsBlockMutation(string stream, int count, int first) {
        string source = Encoding.ASCII.GetString(BuildOpaqueMarkerPdf("Safe metadata", inStream: false));
        string compressed = "6 0 obj\n<< /Type /ObjStm /N " + count + " /First " + first +
            " /Length " + stream.Length + " >>\nstream\n" + stream + "\nendstream\nendobj\n";
        byte[] pdf = Encoding.ASCII.GetBytes(source.Replace("trailer", compressed + "trailer"));
        var preflight = PdfInspector.Preflight(pdf);
        Assert.Contains(preflight.RewriteBlockers, blocker => blocker.Kind == PdfRewriteBlockerKind.IncompleteObjectGraph);
        foreach (PdfMutationOperation operation in System.Enum.GetValues(typeof(PdfMutationOperation)))
            Assert.False(PdfMutationPlanner.Plan(preflight, operation).CanExecute);
        Assert.ThrowsAny<System.Exception>(() => PdfDocument.Load(pdf).Security.Encrypt(new("open") { OwnerPassword = "owner" }));
    }

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
