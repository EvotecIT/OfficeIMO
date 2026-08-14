using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void MalformedPngXmpChunkMakesTheValidCarrierAmbiguous() {
        byte[] prefix = Join(Encoding.ASCII.GetBytes("XML:com.adobe.xmp"), new byte[] { 0, 0, 0, 0, 0 });
        byte[] malformed = Join(Encoding.ASCII.GetBytes("XML:com.adobe.xmp"), new byte[] { 0, 1 });
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[13]),
            CreatePngChunk("iTXt", Join(prefix, CreateXmpPacket())),
            CreatePngChunk("iTXt", malformed),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void StructuredManifestSkipsEmbeddedEndDelimiterText() {
        string inputText = "-----BEGIN C2PA MANIFEST-----\n" +
            "https://example.test/path-----END C2PA MANIFEST-----\n" +
            "-----END C2PA MANIFEST-----\n";
        byte[] input = Encoding.UTF8.GetBytes(inputText);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.txt");

        Assert.True(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void OpcSignatureOriginTextInsideCommentsDoesNotBlockMutation() {
        const string relationships =
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            "<!-- http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin -->" +
            "</Relationships>";
        byte[] package = CreateZip(
            ("_rels/.rels", Encoding.UTF8.GetBytes(relationships)),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "publication.zip");

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void GifTrailerMustTerminateTheAssetBeforeStrictRemoval() {
        byte[] gif = Join(
            Encoding.ASCII.GetBytes("GIF89a"),
            new byte[7],
            CreateGifApplication("C2PA_GIF", new byte[] { 1, 0, 0 }, CreateManifestStore()),
            new byte[] { 0x3B, 0x00 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(gif, "fixture.gif");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(gif, result.ToArray());
    }

    [Fact]
    public void DuplicateWebpExtendedHeadersInvalidateEarlierXmp() {
        byte[] webp = CreateWebp(
            CreateVp8xChunk(advertiseXmp: true),
            CreateRiffChunk("VP8 ", new byte[] { 1, 2 }),
            CreateRiffChunk("XMP ", CreateXmpPacket()),
            CreateVp8xChunk(advertiseXmp: true));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(webp, result.ToArray());
    }

    [Fact]
    public void ManifestRejectsDataBoxStoreBeforeRequiredChildren() {
        byte[] description = CreateBox("jumd", Join(
            C2paUuid("c2db"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.databoxes\0")));
        byte[] dataBoxes = CreateBox("jumb", Join(description, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] manifest = InsertFirstManifestChild(CreateManifestStore(), dataBoxes);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(
            CreatePngWithC2paManifest(manifest), "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void TiffRejectsShortTypedSubIfdPointers() {
        const int primaryIfdOffset = 8;
        const int subIfdOffset = 26;
        byte[] tiff = new byte[subIfdOffset + 6];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = primaryIfdOffset;
        tiff[primaryIfdOffset] = 1;
        WriteLittleEndianEntry(tiff, primaryIfdOffset + 2, 330, 3, 1, subIfdOffset);

        Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceInspector.Inspect(tiff, "fixture.tif"));
    }

    private static byte[] InsertFirstManifestChild(byte[] store, byte[] child) {
        int storeDescriptionLength = ReadBigEndianInt32(store, 8);
        int manifestOffset = 8 + storeDescriptionLength;
        int manifestDescriptionLength = ReadBigEndianInt32(store, manifestOffset + 8);
        int insertionOffset = manifestOffset + 8 + manifestDescriptionLength;
        byte[] result = new byte[store.Length + child.Length];
        Buffer.BlockCopy(store, 0, result, 0, insertionOffset);
        Buffer.BlockCopy(child, 0, result, insertionOffset, child.Length);
        Buffer.BlockCopy(store, insertionOffset, result, insertionOffset + child.Length, store.Length - insertionOffset);
        WriteBigEndian(result, manifestOffset, ReadBigEndianInt32(store, manifestOffset) + child.Length);
        WriteBigEndian(result, 0, result.Length);
        return result;
    }
}
