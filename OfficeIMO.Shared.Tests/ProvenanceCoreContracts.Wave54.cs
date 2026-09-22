using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void SvgManifestTextIsRejectedAtTheConfiguredPacketLimit() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:c2pa=\"http://c2pa.org/manifest\">" +
            "<metadata><c2pa:manifest>" + new string('A', 1024 * 1024) +
            "</c2pa:manifest></metadata></svg>");
        var options = new OfficeProvenanceOptions {
            MaxAssetBytes = 2 * 1024 * 1024,
            MaxManifestBytes = 1024
        };

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(svg, "fixture.svg", options));
    }

    [Fact]
    public void GifRawXmpDoesNotChargeTentativeSubBlockProbe() {
        byte[] gif = Join(Encoding.ASCII.GetBytes("GIF89a"),
            new byte[] { 1, 0, 1, 0, 0, 0, 0 },
            CreateGifXmpExtension(Join(CreateXmpPacket(), Encoding.ASCII.GetBytes(new string(' ', 2048)))),
            CreateMinimalGifImage(), new byte[] { 0x3B });
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 32 };

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(gif, "fixture.gif", options);

        Assert.Contains(report.Evidence, evidence =>
            evidence.Carrier == OfficeProvenanceCarrierKind.IptcDigitalSourceType);
    }

    [Fact]
    public void GifRawXmpRemainsRawWhenItsBytesAlignWithSubBlockLengths() {
        byte[] xmp = CreateXmpPacket();
        int cursor = 0;
        while (cursor < xmp.Length) cursor += xmp[cursor] + 1;
        byte[] alignedXmp = Join(xmp, Encoding.ASCII.GetBytes(new string(' ', cursor - xmp.Length)));
        byte[] gif = Join(Encoding.ASCII.GetBytes("GIF89a"),
            new byte[] { 1, 0, 1, 0, 0, 0, 0 },
            CreateGifXmpExtension(alignedXmp), CreateMinimalGifImage(), new byte[] { 0x3B });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(gif, "fixture.gif");

        Assert.Contains(report.Evidence, evidence =>
            evidence.Carrier == OfficeProvenanceCarrierKind.IptcDigitalSourceType);
    }

    [Fact]
    public void GifRawXmpDisambiguationHonorsTheXmlNodeLimit() {
        byte[] xmp = Encoding.UTF8.GetBytes("<root>" + string.Concat(Enumerable.Repeat("<item/>", 32)) + "</root>");
        int cursor = 0;
        while (cursor < xmp.Length) cursor += xmp[cursor] + 1;
        byte[] alignedXmp = Join(xmp, Encoding.ASCII.GetBytes(new string(' ', cursor - xmp.Length)));
        byte[] gif = Join(Encoding.ASCII.GetBytes("GIF89a"),
            new byte[] { 1, 0, 1, 0, 0, 0, 0 },
            CreateGifXmpExtension(alignedXmp), CreateMinimalGifImage(), new byte[] { 0x3B });

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(
            gif, "fixture.gif", new OfficeProvenanceOptions { MaxContainerEntries = 10 }));

        Assert.Contains("XML node limit", exception.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void TiffStripValidationChecksExactInflatedBytesAndChecksum() {
        byte[] pixels = new byte[64 * 1024];
        byte[] compressed = OfficeZlibCodec.Compress(pixels);
        Assert.True(OfficeTiffCodec.TryValidateStripPayload(
            compressed, 0, compressed.Length, (int)OfficeTiffCompression.Deflate, pixels.Length));

        compressed[compressed.Length - 1] ^= 1;
        Assert.False(OfficeTiffCodec.TryValidateStripPayload(
            compressed, 0, compressed.Length, (int)OfficeTiffCompression.Deflate, pixels.Length));
        Assert.False(OfficeTiffCodec.TryValidateStripPayload(
            new byte[] { 0xFF, 0 }, 0, 2, (int)OfficeTiffCompression.PackBits, 3));
    }

    [Fact]
    public void PngXmpPacketRespectsTheProvenancePacketLimit() {
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[] { 0, 0, 0, 1, 0, 0, 0, 1, 8, 2, 0, 0, 0 }),
            CreatePngChunk("iTXt", Join(
                Encoding.ASCII.GetBytes("XML:com.adobe.xmp"),
                new byte[] { 0, 0, 0, 0, 0 },
                Encoding.UTF8.GetBytes(new string('x', 4096)))),
            CreatePngChunk("IDAT", CreateValidPngImageData()),
            CreatePngChunk("IEND", Array.Empty<byte>()));
        var options = new OfficeProvenanceOptions { MaxManifestBytes = 1024 };

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(png, "fixture.png", options));
    }

    [Fact]
    public void GifXmpSubBlocksStartingWithXmlLikeLengthStillConsumeEntryBudget() {
        byte[] firstBlock = new byte[60];
        firstBlock[0] = (byte)'?';
        byte[] extension = CreateGifXmpExtension(Join(
            new byte[] { 0x3C }, firstBlock,
            new byte[] { 1, (byte)'x' }));
        byte[] gif = Join(Encoding.ASCII.GetBytes("GIF89a"),
            new byte[] { 1, 0, 1, 0, 0, 0, 0 },
            extension, CreateMinimalGifImage(), new byte[] { 0x3B });
        var options = new OfficeProvenanceOptions { MaxContainerEntries = 2 };

        Assert.Throws<InvalidDataException>(() => OfficeProvenanceInspector.Inspect(gif, "fixture.gif", options));
    }

    [Fact]
    public void SvgMetadataRootDirectIptcAttributeIsRemoved() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\">" +
            "<metadata iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Single(result.Before.Evidence);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain("trainedAlgorithmicMedia", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void SvgMetadataRootAndNestedXmpAreRemovedAsOneCarrier(bool requireValidCarrier) {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:iptc='http://iptc.org/std/Iptc4xmpExt/2008-02-29/' " +
            "xmlns:x='adobe:ns:meta/' xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'>" +
            "<metadata iptc:DigitalSourceType='trainedAlgorithmicMedia'><x:xmpmeta><rdf:RDF>" +
            "<rdf:Description iptc:DigitalSourceType='trainedAlgorithmicMedia'/></rdf:RDF></x:xmpmeta>" +
            "</metadata></svg>");
        var options = new OfficeProvenanceRemovalOptions {
            RequireStructurallyValidCarrier = requireValidCarrier
        };

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg", options);

        Assert.True(result.WasChanged);
        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain("trainedAlgorithmicMedia", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void SvgNestedManifestIsRemovedBeforeReplacingItsMetadataScope() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:iptc='http://iptc.org/std/Iptc4xmpExt/2008-02-29/' " +
            "xmlns:c2pa='http://c2pa.org/manifest'><metadata iptc:DigitalSourceType='trainedAlgorithmicMedia'>" +
            "<c2pa:manifest>" + manifest + "</c2pa:manifest></metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.True(result.WasChanged);
        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.Empty(result.After.Evidence);
        Assert.DoesNotContain("trainedAlgorithmicMedia", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
        Assert.DoesNotContain("c2pa:manifest", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void SvgNestedDuplicateXmpPacketsRemainStructurallyInvalid() {
        const string packet = "<x:xmpmeta><rdf:RDF><rdf:Description " +
            "iptc:DigitalSourceType='trainedAlgorithmicMedia'/></rdf:RDF></x:xmpmeta>";
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:iptc='http://iptc.org/std/Iptc4xmpExt/2008-02-29/' " +
            "xmlns:x='adobe:ns:meta/' xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'>" +
            "<metadata iptc:DigitalSourceType='trainedAlgorithmicMedia'>" + packet + packet +
            "</metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Equal(3, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void SvgMetadataDeclarationDoesNotHideMultipleRdfScopes() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:iptc='http://iptc.org/std/Iptc4xmpExt/2008-02-29/' " +
            "xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#'><metadata " +
            "iptc:DigitalSourceType='trainedAlgorithmicMedia'>" +
            "<rdf:RDF><rdf:Description iptc:DigitalSourceType='trainedAlgorithmicMedia'/></rdf:RDF>" +
            "<rdf:RDF><rdf:Description iptc:DigitalSourceType='trainedAlgorithmicMedia'/></rdf:RDF>" +
            "</metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Equal(3, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(svg, result.ToArray());
    }

    [Fact]
    public void DuplicateWebpExtendedHeadersInvalidateC2pa() {
        byte[] webp = CreateWebp(
            CreateVp8xChunk(advertiseXmp: false),
            CreateRiffChunk("VP8 ", new byte[] { 1, 2 }),
            CreateVp8xChunk(advertiseXmp: false),
            CreateRiffChunk("C2PA", CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(webp, "fixture.webp");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(webp, result.ToArray());
    }

    [Fact]
    public void NoncontiguousPngImageDataInvalidatesC2pa() {
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[13]),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IDAT", new byte[] { 1 }),
            CreatePngChunk("tEXt", Encoding.ASCII.GetBytes("separator")),
            CreatePngChunk("IDAT", new byte[] { 2 }),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void MultipleDirectSvgRdfScopesAreStructurallyAmbiguous() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><metadata>" +
            "<rdf:RDF><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></rdf:RDF>" +
            "<rdf:RDF><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></rdf:RDF>" +
            "</metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(svg, result.ToArray());
    }
}
