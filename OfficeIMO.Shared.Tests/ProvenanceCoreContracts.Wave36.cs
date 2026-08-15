using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void DuplicateSvgManifestElementsAreAllStructurallyInvalid() {
        string encoded = Convert.ToBase64String(CreateManifestStore());
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:c2pa=\"http://c2pa.org/manifest\"><metadata>" +
            $"<c2pa:manifest>{encoded}</c2pa:manifest><c2pa:manifest>{encoded}</c2pa:manifest>" +
            "</metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(svg, result.ToArray());
    }

    [Fact]
    public void DuplicateManifestLabelsInvalidateTheManifestStore() {
        byte[] store = CreateManifestStore();
        int descriptionLength = ReadBigEndianInt32(store, 8);
        int manifestOffset = 8 + descriptionLength;
        int manifestLength = ReadBigEndianInt32(store, manifestOffset);
        byte[] duplicatedPayload = new byte[descriptionLength + manifestLength * 2];
        Buffer.BlockCopy(store, 8, duplicatedPayload, 0, descriptionLength);
        Buffer.BlockCopy(store, manifestOffset, duplicatedPayload, descriptionLength, manifestLength);
        Buffer.BlockCopy(store, manifestOffset, duplicatedPayload, descriptionLength + manifestLength, manifestLength);
        byte[] duplicatedStore = CreateBox("jumb", duplicatedPayload);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(
            CreatePngWithC2paManifest(duplicatedStore), "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void LaterDuplicatePngHeaderInvalidatesEarlierManifestCarrier() {
        byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
        byte[] header = CreatePngChunk("IHDR", new byte[13]);
        byte[] png = Join(
            signature,
            header,
            CreatePngChunk("caBX", CreateManifestStore()),
            header,
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }
}
