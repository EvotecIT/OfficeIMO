using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceReviewRegressionContracts {
    [Fact]
    public void JpegExtendedXmpIgnoresNonScalarDigestReferences() {
        byte[] extendedPacket = Encoding.UTF8.GetBytes(
            "<rdf:RDF xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><rdf:Description " +
            "iptc:DigitalSourceType=\"http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia\"/></rdf:RDF>");
        string guid = ComputeMd5(extendedPacket);
        byte[] standardPacket = Encoding.UTF8.GetBytes(
            $"<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:xmpNote=\"http://ns.adobe.com/xmp/note/\" " +
            $"xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\"><xmpNote:HasExtendedXMP>" +
            $"<rdf:value>{guid}</rdf:value></xmpNote:HasExtendedXMP></x:xmpmeta>");
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegSegment(0xE1, Join(Encoding.ASCII.GetBytes("http://ns.adobe.com/xap/1.0/\0"), standardPacket)),
            CreateExtendedXmpSegment(guid, extendedPacket),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(jpeg, "fixture.jpg");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }
}
