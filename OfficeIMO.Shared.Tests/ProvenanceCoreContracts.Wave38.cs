using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void UnicodePathSignatureEntryBlocksGenericZipMutation() {
        byte[] package = CreateZipWithUnicodePathEntry(
            "META-INF/harmless.xml",
            "META-INF/documentsignatures.xml",
            Encoding.UTF8.GetBytes("<signature/>"));

        Assert.Throws<InvalidOperationException>(() =>
            OfficeProvenanceRemover.Remove(package, "publication.zip"));
    }

    [Fact]
    public void DuplicateAssertionLabelsInvalidateTheManifest() {
        byte[] manifest = DuplicateFirstAssertion(CreateManifestStore());
        byte[] png = CreatePngWithC2paManifest(manifest);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void DuplicateDataBoxStoresInvalidateTheManifest() {
        byte[] description = CreateBox("jumd", Join(
            C2paUuid("c2db"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.databoxes\0")));
        byte[] dataBoxes = CreateBox("jumb", Join(description, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] manifest = AppendDirectManifestChild(
            AppendDirectManifestChild(CreateManifestStore(), dataBoxes),
            dataBoxes);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(
            CreatePngWithC2paManifest(manifest), "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void DuplicatePngXmpChunksAreStructurallyInvalid() {
        byte[] prefix = Join(Encoding.ASCII.GetBytes("XML:com.adobe.xmp"), new byte[] { 0, 0, 0, 0, 0 });
        byte[] xmp = CreatePngChunk("iTXt", Join(prefix, CreateXmpPacket()));
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[13]),
            xmp,
            xmp,
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void DuplicateGifXmpExtensionsAreStructurallyInvalid() {
        byte[] xmp = CreateGifXmpExtension(CreateXmpPacket());
        byte[] gif = Join(
            Encoding.ASCII.GetBytes("GIF89a"),
            new byte[7],
            xmp,
            xmp,
            new byte[] { 0x3B });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(gif, "fixture.gif");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(gif, result.ToArray());
    }

    [Fact]
    public void DuplicateTiffXmpTagsAreStructurallyInvalid() {
        byte[] xmp = CreateXmpPacket();
        const int firstPayloadOffset = 38;
        int secondPayloadOffset = firstPayloadOffset + xmp.Length;
        byte[] tiff = new byte[secondPayloadOffset + xmp.Length];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 2;
        WriteLittleEndianEntry(tiff, 10, 700, 1, xmp.Length, firstPayloadOffset);
        WriteLittleEndianEntry(tiff, 22, 700, 1, xmp.Length, secondPayloadOffset);
        Buffer.BlockCopy(xmp, 0, tiff, firstPayloadOffset, xmp.Length);
        Buffer.BlockCopy(xmp, 0, tiff, secondPayloadOffset, xmp.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(tiff, result.ToArray());
    }

    private static byte[] DuplicateFirstAssertion(byte[] store) {
        int storeDescriptionLength = ReadBigEndianInt32(store, 8);
        int manifestOffset = 8 + storeDescriptionLength;
        int manifestDescriptionLength = ReadBigEndianInt32(store, manifestOffset + 8);
        int assertionStoreOffset = manifestOffset + 8 + manifestDescriptionLength;
        int assertionStoreDescriptionLength = ReadBigEndianInt32(store, assertionStoreOffset + 8);
        int assertionOffset = assertionStoreOffset + 8 + assertionStoreDescriptionLength;
        int assertionLength = ReadBigEndianInt32(store, assertionOffset);
        int insertionOffset = assertionOffset + assertionLength;

        byte[] result = new byte[store.Length + assertionLength];
        Buffer.BlockCopy(store, 0, result, 0, insertionOffset);
        Buffer.BlockCopy(store, assertionOffset, result, insertionOffset, assertionLength);
        Buffer.BlockCopy(store, insertionOffset, result, insertionOffset + assertionLength, store.Length - insertionOffset);
        WriteBigEndian(result, assertionStoreOffset, ReadBigEndianInt32(store, assertionStoreOffset) + assertionLength);
        WriteBigEndian(result, manifestOffset, ReadBigEndianInt32(store, manifestOffset) + assertionLength);
        WriteBigEndian(result, 0, result.Length);
        return result;
    }

    private static byte[] CreateGifXmpExtension(byte[] packet) {
        byte[] trailer = new byte[258];
        trailer[0] = 0x01;
        for (int index = 1; index <= 255; index++) trailer[index] = checked((byte)(256 - index));
        return Join(
            new byte[] { 0x21, 0xFF, 0x0B },
            Encoding.ASCII.GetBytes("XMP DataXMP"),
            packet,
            trailer);
    }
}
