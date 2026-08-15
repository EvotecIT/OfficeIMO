using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void StrictRemovalPreservesDuplicateStructuredTextAndSelectorWrappers() {
        string block = "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n";
        byte[] structured = Encoding.UTF8.GetBytes(block + block);
        byte[] wrapper = CreateTextWrapper(CreateManifestStore());
        byte[] selectors = Join(wrapper, wrapper);

        OfficeProvenanceRemovalResult structuredResult = OfficeProvenanceRemover.Remove(structured, "fixture.md");
        OfficeProvenanceRemovalResult selectorResult = OfficeProvenanceRemover.Remove(selectors, "fixture.txt");

        Assert.Equal(2, structuredResult.Before.Evidence.Count);
        Assert.All(structuredResult.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.False(structuredResult.WasChanged);
        Assert.Equal(structured, structuredResult.ToArray());
        Assert.Equal(2, selectorResult.Before.Evidence.Count);
        Assert.All(selectorResult.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.False(selectorResult.WasChanged);
        Assert.Equal(selectors, selectorResult.ToArray());
    }

    [Fact]
    public void StrictRemovalPreservesDuplicateStandardJpegXmpPackets() {
        byte[] header = Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0");
        byte[] segment = CreateJpegSegment(0xE1, Join(header, CreateXmpPacket()));
        byte[] jpeg = Join(new byte[] { 0xFF, 0xD8 }, segment, segment, new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.True(result.Before.HasGenerativeAiDeclaration);
        Assert.All(result.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }

    [Fact]
    public void StrictRemovalPreservesNonUtf8PngITXtXmp() {
        byte[] prefix = Encoding.ASCII.GetBytes("XML:com.adobe.xmp\0\0\0\0\0");
        byte[] utf16 = Join(new byte[] { 0xFF, 0xFE }, Encoding.Unicode.GetBytes(Encoding.UTF8.GetString(CreateXmpPacket())));
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 }),
            CreatePngChunk("iTXt", Join(prefix, utf16)),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.True(result.Before.HasGenerativeAiDeclaration);
        Assert.All(result.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void RemovalHonorsTheNestedEmbeddedAssetDisableSwitch() {
        byte[] package = CreateZip(("media/image.png", Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 }),
            CreatePngChunk("iTXt", Join(Encoding.ASCII.GetBytes("XML:com.adobe.xmp\0\0\0\0\0"), CreateXmpPacket())),
            CreatePngChunk("IEND", Array.Empty<byte>()))));
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.ProcessEmbeddedAssets = false;

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "fixture.zip", options);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(package, result.ToArray());
    }
}
