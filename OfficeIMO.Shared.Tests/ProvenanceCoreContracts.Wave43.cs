using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void RdfXmlLiteralAttributesAreNotTreatedAsDigitalSourceMetadata() {
        byte[] packet = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><rdf:RDF><rdf:Description>" +
            "<x:Unrelated rdf:parseType=\"Literal\"><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></x:Unrelated>" +
            "</rdf:Description></rdf:RDF></x:xmpmeta>");
        byte[] prefix = Join(Encoding.ASCII.GetBytes("XML:com.adobe.xmp"), new byte[] { 0, 0, 0, 0, 0 });
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[13]),
            CreatePngChunk("iTXt", Join(prefix, packet)),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(png, "fixture.png");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void UnsortedTiffXmpDirectoryIsStructurallyInvalid() {
        byte[] xmp = CreateXmpPacket();
        const int payloadOffset = 38;
        byte[] tiff = new byte[payloadOffset + xmp.Length];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 2;
        WriteLittleEndianEntry(tiff, 10, 700, 1, xmp.Length, payloadOffset);
        WriteLittleEndianEntry(tiff, 22, 256, 3, 1, 1);
        Buffer.BlockCopy(xmp, 0, tiff, payloadOffset, xmp.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(tiff, result.ToArray());
    }

    [Fact]
    public void UndefinedTiffXmpTagIsStructurallyInvalid() {
        byte[] xmp = CreateXmpPacket();
        const int payloadOffset = 26;
        byte[] tiff = new byte[payloadOffset + xmp.Length];
        tiff[0] = tiff[1] = (byte)'I';
        tiff[2] = 42;
        tiff[4] = 8;
        tiff[8] = 1;
        WriteLittleEndianEntry(tiff, 10, 700, 7, xmp.Length, payloadOffset);
        Buffer.BlockCopy(xmp, 0, tiff, payloadOffset, xmp.Length);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(tiff, "fixture.tif");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(tiff, result.ToArray());
    }

    [Fact]
    public void DuplicateSvgXmpPacketsAreStructurallyInvalid() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:x=\"adobe:ns:meta/\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><metadata>" +
            "<x:xmpmeta><rdf:RDF><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta>" +
            "<x:xmpmeta><rdf:RDF><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta>" +
            "</metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(svg, result.ToArray());
    }
}
