using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void SvgReportsAndRemovesMixedCarriersInDocumentOrder() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string svg = "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:c2pa=\"http://c2pa.org/manifest\"><metadata>" +
            Encoding.UTF8.GetString(CreateXmpPacket()) +
            $"<c2pa:manifest>{manifest}</c2pa:manifest></metadata></svg>";
        byte[] input = Encoding.UTF8.GetBytes(svg);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(input, "fixture.svg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.svg");

        Assert.Equal(OfficeProvenanceCarrierKind.IptcDigitalSourceType, report.Evidence[0].Carrier);
        Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, report.Evidence[report.Evidence.Count - 1].Carrier);
        Assert.Equal(OfficeProvenanceCarrierKind.IptcDigitalSourceType, result.Changes[0].Carrier);
        Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, result.Changes[result.Changes.Count - 1].Carrier);
    }

    [Fact]
    public void JpegOrdersTrailingSuffixResultsBySourceOffset() {
        byte[] manifest = CreateManifestStore();
        byte[] xmpHeader = Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegApp11(manifest, 0, manifest.Length, 1, 1),
            CreateJpegSegment(0xE1, Join(xmpHeader, CreateXmpPacket())),
            CreateMinimalJpegFrame(),
            CreateMinimalJpegScan(),
            new byte[] { 0, 0xFF, 0xD9, 0xDE, 0xAD });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, report.Evidence[0].Carrier);
        Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, result.Changes[0].Carrier);
        byte[] output = result.ToArray();
        Assert.Equal(new byte[] { 0xDE, 0xAD }, output.Skip(output.Length - 2).ToArray());
    }

    [Fact]
    public void DuplicatePrimaryTiffC2paTagsAreStructurallyInvalid() {
        byte[] manifest = CreateManifestStore();
        const int payloadOffset = 38;
        byte[] tiff = new byte[payloadOffset + (manifest.Length * 2)];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 2;
        WriteLittleEndianEntry(tiff, 10, 0xCD41, 7, manifest.Length, payloadOffset);
        WriteLittleEndianEntry(tiff, 22, 0xCD41, 7, manifest.Length, payloadOffset + manifest.Length);
        Buffer.BlockCopy(manifest, 0, tiff, payloadOffset, manifest.Length);
        Buffer.BlockCopy(manifest, 0, tiff, payloadOffset + manifest.Length, manifest.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void PngXmpBeforeIhdrIsStructurallyInvalid() {
        byte[] header = { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A };
        byte[] prefix = Join(Encoding.ASCII.GetBytes("XML:com.adobe.xmp"), new byte[] { 0, 0, 0, 0, 0 });
        byte[] ihdr = { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 };
        byte[] png = Join(header,
            CreatePngChunk("iTXt", Join(prefix, CreateXmpPacket())),
            CreatePngChunk("IHDR", ihdr),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void AssertionSuperboxRejectsArbitraryContentBoxes() {
        byte[] manifest = CreateManifestStore();
        int contentType = FindAscii(manifest, "cbor");
        Assert.True(contentType >= 0);
        Encoding.ASCII.GetBytes("free").CopyTo(manifest, contentType);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(
            CreatePngWithC2paManifest(manifest), "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    private static int FindAscii(byte[] data, string value) {
        byte[] expected = Encoding.ASCII.GetBytes(value);
        for (int offset = 0; offset <= data.Length - expected.Length; offset++) {
            if (data.AsSpan(offset, expected.Length).SequenceEqual(expected)) return offset;
        }
        return -1;
    }
}
