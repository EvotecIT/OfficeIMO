using System.IO.Compression;
using System.Text;
using OfficeIMO;
using OfficeIMO.Provenance;
using OfficeIMO.Security;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void LaterInvalidWebpPaddingPreservesEarlierXmpCarrier() {
        byte[] badTrailingChunk = CreateRiffChunk("EXIF", new byte[] { 1 });
        badTrailingChunk[badTrailingChunk.Length - 1] = 0x7f;
        byte[] webp = CreateWebp(
            CreateVp8xChunk(advertiseXmp: true),
            CreateRiffChunk("VP8 ", new byte[] { 1, 2 }),
            CreateRiffChunk("XMP ", CreateXmpPacket()),
            badTrailingChunk);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(webp, result.ToArray());
    }

    [Fact]
    public void ConventionalApplicationSignatureMetadataSurvivesMissingContentTypeDeclaration() {
        byte[] package = CreateZip(
            ("docProps/app.xml", Encoding.UTF8.GetBytes(
                "<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">" +
                "<DigSig>signature</DigSig></Properties>")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        OfficePackageSignatureInfo info = OfficePackageSignatureService.Inspect(package);

        Assert.True(info.HasApplicationSignatureMetadata);
        Assert.True(info.HasSignatures);
    }
}
