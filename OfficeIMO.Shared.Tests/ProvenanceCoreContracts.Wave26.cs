using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PngC2paRequiresOneValidFirstIhdr(bool duplicateHeader) {
        byte[] header = CreatePngChunk("IHDR", new byte[13]);
        byte[] malformedPrefix;
        if (duplicateHeader) {
            malformedPrefix = Join(header, CreatePngChunk("IHDR", new byte[13]));
        } else {
            header[header.Length - 1] ^= 0x01;
            malformedPrefix = header;
        }
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            malformedPrefix,
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void OpcSignaturePartsRemainEvidenceWithoutContentTypes() {
        byte[] package = CreateZip(
            ("_xmlsignatures/sig1.xml", Encoding.UTF8.GetBytes("<Signature/>")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            OfficeProvenanceRemover.Remove(package, "document.docx"));

        Assert.Contains("invalidate package signatures", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OpcSignatureOriginRelationshipsRemainEvidenceWithoutContentTypes() {
        byte[] package = CreateZip(
            ("_rels/.rels", Encoding.UTF8.GetBytes(
                "<Relationships><Relationship Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"missing.sigs\"/></Relationships>")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            OfficeProvenanceRemover.Remove(package, "document.docx"));

        Assert.Contains("invalidate package signatures", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
