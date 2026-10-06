using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void JpegRepeatedContinuationHeadersMustMatchTheInitialHeader(bool corruptHeader) {
        byte[] manifest = CreateManifestStore(324);
        byte[] first = CreateJpegApp11(manifest, 0, 160, 1, 1);
        byte[] tail = Join(manifest.Take(8).ToArray(), manifest.Skip(160).ToArray());
        if (corruptHeader) tail[3] ^= 1;
        byte[] second = CreateJpegApp11(tail, 0, tail.Length, 1, 2);
        var result = OfficeProvenanceRemover.Remove(CreateValidJpeg(first, second), "fragmented.jpg");
        Assert.Equal(!corruptHeader, Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.Equal(!corruptHeader, result.WasChanged);
        Assert.Equal(corruptHeader, result.After.HasC2paManifest);
    }

    [Theory]
    [InlineData(16, true)]
    [InlineData(32, true)]
    [InlineData(0, false)]
    [InlineData(15, false)]
    [InlineData(31, false)]
    [InlineData(33, false)]
    public void SaltedAssertionDescriptionsRequireTheSpecifiedSaltSize(int length, bool valid) {
        byte[] input = CreatePngWithC2paManifest(CreateManifestStore(assertionSalt: CreateBox("c2sh", new byte[length])));
        var result = OfficeProvenanceRemover.Remove(input, "salted.png");
        Assert.Equal(valid, Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.Equal(valid, result.WasChanged);
        if (!valid) Assert.Equal(input, result.ToArray());
    }

    [Theory]
    [InlineData("pair", true)]
    [InlineData("descriptor", false)]
    [InlineData("data", false)]
    [InlineData("empty", false)]
    [InlineData("trailing", false)]
    public void EmbeddedFileAssertionsRequireACompleteBoundedPair(string variant, bool valid) {
        byte[] description = CreateBox("bfdb", Join(new byte[] { 0 }, Encoding.ASCII.GetBytes("image/jpeg\0")));
        byte[] data = CreateBox("bidb", new byte[] { 1, 2, 3 });
        byte[] content = variant switch {
            "descriptor" => description,
            "data" => data,
            "empty" => Join(description, CreateBox("bidb", Array.Empty<byte>())),
            "trailing" => Join(description, data, CreateBox("free", new byte[] { 0 })),
            _ => Join(description, data)
        };
        byte[] input = CreatePngWithC2paManifest(CreateManifestStore(assertionContent: content));
        var result = OfficeProvenanceRemover.Remove(input, "thumbnail.png");
        Assert.Equal(valid, Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.Equal(valid, result.WasChanged);
        if (!valid) Assert.Equal(input, result.ToArray());
    }
}
