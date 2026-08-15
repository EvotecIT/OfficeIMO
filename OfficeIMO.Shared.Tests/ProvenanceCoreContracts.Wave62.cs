using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void IndexedPngWithoutRequiredPalettePreservesProvenanceCarriers() {
        byte[] indexedHeader = { 0, 0, 0, 1, 0, 0, 0, 1, 8, 3, 0, 0, 0 };
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", indexedHeader),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IDAT", new byte[] { 1 }),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void IndexedPngPaletteMustPrecedeImageData() {
        byte[] indexedHeader = { 0, 0, 0, 1, 0, 0, 0, 1, 8, 3, 0, 0, 0 };
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", indexedHeader),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IDAT", new byte[] { 1 }),
            CreatePngChunk("PLTE", new byte[] { 0, 0, 0 }),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void IndexedPngWithAValidPaletteRemovesProvenanceCarriers() {
        byte[] indexedHeader = { 0, 0, 0, 1, 0, 0, 0, 1, 8, 3, 0, 0, 0 };
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", indexedHeader),
            CreatePngChunk("PLTE", new byte[] { 0, 0, 0 }),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IDAT", new byte[] { 1 }),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.True(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }
}
