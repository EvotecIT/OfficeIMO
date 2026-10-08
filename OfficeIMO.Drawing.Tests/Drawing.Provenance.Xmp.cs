using System.IO;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class XmpProvenanceInspectionTests {
    [Theory]
    [InlineData("trainedAlgorithmicMedia", OfficeProvenanceDigitalSourceKind.TrainedAlgorithmicMedia)]
    [InlineData("compositeWithTrainedAlgorithmicMedia", OfficeProvenanceDigitalSourceKind.CompositeWithTrainedAlgorithmicMedia)]
    public void IndependentXmpPacketPreservesStandardizedSourceDeclarations(string term, OfficeProvenanceDigitalSourceKind expected) {
        string value = "http://cv.iptc.org/newscodes/digitalsourcetype/" + term;
        byte[] packet = Encoding.UTF8.GetBytes("<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' " +
            "xmlns:iptc='http://iptc.org/std/Iptc4xmpExt/2008-02-29/'><rdf:Description><iptc:DigitalSourceType rdf:resource='" + value + "'/></rdf:Description></rdf:RDF>");

        OfficeProvenanceReport report = OfficeProvenanceInspector.InspectXmp(packet);

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.Equal(OfficeProvenanceCarrierKind.IptcDigitalSourceType, evidence.Carrier);
        Assert.Equal(expected, evidence.DigitalSourceKind);
        Assert.Equal(value, evidence.Value);
        Assert.True(evidence.IsStructurallyValid);
        Assert.True(report.HasGenerativeAiDeclaration);
        Assert.False(report.HasC2paManifest);
    }

    [Theory]
    [InlineData("<rdf:RDF xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' xmlns:decoy='urn:decoy'><rdf:Description decoy:DigitalSourceType='http://cv.iptc.org/newscodes/digitalsourcetype/trainedAlgorithmicMedia'/></rdf:RDF>")]
    [InlineData("<rdf:RDF")]
    public void IndependentXmpPacketDoesNotTreatDecoysOrMalformedXmlAsDeclarations(string xml) {
        OfficeProvenanceReport report = OfficeProvenanceInspector.InspectXmp(Encoding.UTF8.GetBytes(xml));
        Assert.Empty(report.Evidence);
        Assert.False(report.HasGenerativeAiDeclaration);
    }

    [Fact]
    public void IndependentXmpPacketEnforcesEncodedByteAndXmlNodeLimits() {
        byte[] packet = Encoding.UTF8.GetBytes("<root><one/><two/><three/></root>");
        Assert.ThrowsAny<InvalidDataException>(() => OfficeProvenanceInspector.InspectXmp(packet,
            new OfficeProvenanceOptions { MaxAssetBytes = packet.Length - 1, MaxManifestBytes = packet.Length - 1 }));
        Assert.ThrowsAny<InvalidDataException>(() => OfficeProvenanceInspector.InspectXmp(packet,
            new OfficeProvenanceOptions { MaxContainerEntries = 2 }));
    }
}
