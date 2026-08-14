using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void MixedStructuredAndVariationSelectorTextCarriersAreAmbiguous() {
        string block = "-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n";
        byte[] input = Join(Encoding.UTF8.GetBytes(block), CreateTextWrapper(CreateManifestStore()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(input, "fixture.txt");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(input, result.ToArray());
    }

    [Fact]
    public void MultipleExtendedXmpReferencesInvalidateTheJpegPacketSet() {
        byte[] extendedPacket = Encoding.UTF8.GetBytes(
            "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><rdf:Description " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF>");
        string guid = ComputeMd5(extendedPacket);
        byte[] standardPacket = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:xmpNote=\"http://ns.adobe.com/xmp/note/\">" +
            $"<xmpNote:HasExtendedXMP>{guid}</xmpNote:HasExtendedXMP>" +
            $"<xmpNote:HasExtendedXMP>{guid}</xmpNote:HasExtendedXMP></x:xmpmeta>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), standardPacket)),
            CreateExtendedXmpSegment(guid, extendedPacket),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.NotEmpty(result.Before.Evidence);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }

}
