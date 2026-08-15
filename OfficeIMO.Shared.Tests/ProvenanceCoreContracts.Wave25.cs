using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void MalformedGifXmpProbeDoesNotDoubleChargeSubBlocks() {
        byte[] application = CreateGifApplication("XMP DataXMP", Array.Empty<byte>(), Encoding.ASCII.GetBytes("not-xmp"));
        byte[] gif = Join(Encoding.ASCII.GetBytes("GIF89a"), new byte[7], application, new byte[] { 0x3B });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
            gif, "fixture.gif", new OfficeProvenanceOptions { MaxContainerEntries = 4 });

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void UnsupportedZipEntriesDoNotConsumeTheEmbeddedAssetLimit() {
        byte[] image = CreatePngWithC2paManifest(CreateManifestStore());
        byte[] package = CreateZip(
            ("media/not-an-image.png", Encoding.UTF8.GetBytes("plain text")),
            ("media/credential.png", image));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
            package, "document.docx", new OfficeProvenanceOptions { MaxEmbeddedAssets = 1 });
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(
            package, "document.docx", new OfficeProvenanceRemovalOptions { MaxEmbeddedAssets = 1 });

        Assert.Single(report.Evidence);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void JpegEvidenceAndChangesFollowMarkerOrder() {
        byte[] xmpHeader = Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0");
        byte[] manifest = CreateManifestStore();
        byte[] jpeg = CreateValidJpeg(
            CreateJpegApp11(manifest, 0, manifest.Length, instance: 1, sequence: 1),
            CreateJpegSegment(0xE1, Join(xmpHeader, CreateXmpPacket())));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, report.Evidence[0].Carrier);
        Assert.All(report.Evidence.Skip(1), item => Assert.Equal(OfficeProvenanceCarrierKind.IptcDigitalSourceType, item.Carrier));
        Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, result.Changes[0].Carrier);
    }

    [Fact]
    public void JpegRejectsNestedStartOfImageMarkers() {
        byte[] manifest = CreateManifestStore();
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8, 0xFF, 0xD8 },
            CreateJpegApp11(manifest, 0, manifest.Length, instance: 1, sequence: 1),
            new byte[] { 0xFF, 0xD9 });

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg"));
        Assert.Throws<InvalidDataException>(() => OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg"));
    }

    [Fact]
    public void PngInvalidItextLanguageAndTranslatedKeywordAreStructurallyInvalid() {
        byte[] prefix = Join(
            Encoding.ASCII.GetBytes("XML:com.adobe.xmp"),
            new byte[] { 0, 0, 0, 0xFF, 0, 0xC3, 0x28, 0 });
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[13]),
            CreatePngChunk("iTXt", Join(prefix, CreateXmpPacket())),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.Contains(result.Before.Evidence, item => !item.IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void JumbfDescriptionRejectsUndeclaredTrailingBytes() {
        byte[] store = CreateManifestStore();
        int descriptionLength = ReadBigEndianInt32(store, 8);
        byte[] malformed = new byte[store.Length + 1];
        Buffer.BlockCopy(store, 0, malformed, 0, 8 + descriptionLength);
        malformed[8 + descriptionLength] = 0x7F;
        Buffer.BlockCopy(
            store, 8 + descriptionLength, malformed, 9 + descriptionLength,
            store.Length - 8 - descriptionLength);
        WriteBigEndian(malformed, 8, descriptionLength + 1);
        WriteBigEndian(malformed, 0, malformed.Length);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(
            CreatePngWithC2paManifest(malformed), "fixture.png");

        Assert.False(Assert.Single(report.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void C2paDescriptionsRequireRequestableLabelsAndConsumeDeclaredFields() {
        byte[] missingRequestable = CreateManifestStore();
        missingRequestable[32] = 0x02;
        byte[] withDeclaredFields = AddOuterDescriptionFields(CreateManifestStore());

        OfficeProvenanceReport invalid = OfficeProvenanceInspector.Inspect(
            CreatePngWithC2paManifest(missingRequestable), "fixture.png");
        OfficeProvenanceReport valid = OfficeProvenanceInspector.Inspect(
            CreatePngWithC2paManifest(withDeclaredFields), "fixture.png");

        Assert.False(Assert.Single(invalid.Evidence).IsStructurallyValid);
        Assert.True(Assert.Single(valid.Evidence).IsStructurallyValid);
    }

    [Fact]
    public void SvgManifestElementsRejectNestedMarkup() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        byte[] svg = Encoding.UTF8.GetBytes(
            $"<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:c2pa=\"http://c2pa.org/manifest\"><metadata><c2pa:manifest><span>{manifest}</span></c2pa:manifest></metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void UnknownIptcVocabularyValuesRemainStructurallyValid() {
        byte[] packet = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\"><rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\">" +
            "<rdf:Description xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\" " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/futureValue\"/>" +
            "</rdf:RDF></x:xmpmeta>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), packet)),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceEvidence evidence = Assert.Single(OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg").Evidence);

        Assert.Equal(OfficeProvenanceDigitalSourceKind.Unknown, evidence.DigitalSourceKind);
        Assert.True(evidence.IsStructurallyValid);
    }

    [Fact]
    public void WebpXmpRequiresAValidAdvertisedExtendedFormat() {
        byte[] invalid = CreateWebp(
            CreateValidVp8Chunk(),
            CreateRiffChunk("XMP ", CreateXmpPacket()));
        byte[] valid = CreateWebp(
            CreateVp8xChunk(advertiseXmp: true),
            CreateValidVp8Chunk(),
            CreateRiffChunk("XMP ", CreateXmpPacket()));

        OfficeProvenanceRemovalResult invalidResult = OfficeProvenanceRemover.Remove(invalid, "fixture.webp");
        OfficeProvenanceRemovalResult validResult = OfficeProvenanceRemover.Remove(valid, "fixture.webp");

        Assert.All(invalidResult.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.False(invalidResult.WasChanged);
        Assert.True(validResult.WasChanged);
        Assert.Empty(validResult.After.Evidence.Where(item => item.DigitalSourceKind ==
            OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia));
    }

    private static byte[] CreateVp8xChunk(bool advertiseXmp) {
        byte[] payload = new byte[10];
        if (advertiseXmp) payload[0] = 0x04;
        return CreateRiffChunk("VP8X", payload);
    }

    private static byte[] AddOuterDescriptionFields(byte[] store) {
        int descriptionLength = ReadBigEndianInt32(store, 8);
        const int optionalFieldLength = 4 + 32;
        byte[] result = new byte[store.Length + optionalFieldLength];
        Buffer.BlockCopy(store, 0, result, 0, 8 + descriptionLength);
        Buffer.BlockCopy(
            store, 8 + descriptionLength, result, 8 + descriptionLength + optionalFieldLength,
            store.Length - 8 - descriptionLength);
        result[32] = 0x0F;
        for (int index = 0; index < optionalFieldLength; index++) {
            result[8 + descriptionLength + index] = (byte)(index + 1);
        }
        WriteBigEndian(result, 8, descriptionLength + optionalFieldLength);
        WriteBigEndian(result, 0, result.Length);
        return result;
    }
}
