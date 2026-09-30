using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {


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

    [Theory]
    [MemberData(nameof(DuplicateSvgRdfCarrierScopesAreStructurallyAmbiguousCases))]
    public void DuplicateSvgRdfCarrierScopesAreStructurallyAmbiguous(string caseName, byte[] svg) {
        _ = caseName;
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(svg, result.ToArray());
    }

    public static IEnumerable<object[]> DuplicateSvgRdfCarrierScopesAreStructurallyAmbiguousCases() {
        {
            byte[] svg = Encoding.UTF8.GetBytes(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:x=\"adobe:ns:meta/\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
                "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><metadata>" +
                "<x:xmpmeta><rdf:RDF><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta>" +
                "<x:xmpmeta><rdf:RDF><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></rdf:RDF></x:xmpmeta>" +
                "</metadata></svg>");
            yield return new object[] { "DuplicateSvgXmpPacketsAreStructurallyInvalid", svg };
        }
        {
            byte[] svg = Encoding.UTF8.GetBytes(
                "<svg xmlns=\"http://www.w3.org/2000/svg\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
                "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><metadata>" +
                "<rdf:RDF><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></rdf:RDF>" +
                "<rdf:RDF><rdf:Description iptc:DigitalSourceType=\"trainedAlgorithmicMedia\"/></rdf:RDF>" +
                "</metadata></svg>");
            yield return new object[] { "MultipleDirectSvgRdfScopesAreStructurallyAmbiguous", svg };
        }
    }
}
