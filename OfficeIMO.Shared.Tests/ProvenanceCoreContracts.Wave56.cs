using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void TiffManifestWithoutACompletePrimaryImageIsPreserved() {
        byte[] manifest = CreateManifestStore();
        byte[] tiff = new byte[26 + manifest.Length];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 1;
        WriteLittleEndianEntry(tiff, 10, 0xCD41, 7, manifest.Length, 26);
        Buffer.BlockCopy(manifest, 0, tiff, 26, manifest.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(tiff, result.ToArray());
    }

    [Fact]
    public void OpcSignatureEvidencePathsUseOpcCaseSemantics() {
        byte[] package = CreateZip(
            ("_XMLSIGNATURES/SIG1.XML", Encoding.UTF8.GetBytes("<Signature/>")));

        Assert.True(OfficeProvenanceZip.HasPackageSignature(
            package,
            new OfficeProvenanceRemovalOptions()));
    }
}
