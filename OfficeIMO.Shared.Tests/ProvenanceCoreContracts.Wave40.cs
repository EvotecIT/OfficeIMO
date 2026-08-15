using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void PrivateManifestChildrenCannotReuseReservedLabels() {
        byte[] description = CreateBox("jumd", Join(
            C2paUuid("priv"), new byte[] { 0x03 }, Encoding.ASCII.GetBytes("c2pa.claim\0")));
        byte[] extension = CreateBox("jumb", Join(description, CreateBox("cbor", new byte[] { 0xA0 })));
        byte[] manifest = AppendDirectManifestChild(CreateManifestStore(), extension);

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(
            CreatePngWithC2paManifest(manifest), "fixture.png");

        Assert.False(Assert.Single(result.Before.Evidence).IsStructurallyValid);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void MalformedJpegC2paSequenceInvalidatesACompetingValidSequence() {
        byte[] manifest = CreateManifestStore();
        byte[] malformed = (byte[])manifest.Clone();
        WriteBigEndian(malformed, 0, malformed.Length + 10);
        byte[] jpeg = Join(
            new byte[] { 0xFF, 0xD8 },
            CreateJpegApp11(manifest, 0, manifest.Length, instance: 1, sequence: 1),
            CreateJpegApp11(malformed, 0, malformed.Length, instance: 2, sequence: 1),
            new byte[] { 0xFF, 0xD9 });

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(jpeg, "fixture.jpg");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(jpeg, result.ToArray());
    }

    [Fact]
    public void RdfXmlLiteralMarkupIsNotTreatedAsDigitalSourceMetadata() {
        byte[] packet = Encoding.UTF8.GetBytes(
            "<x:xmpmeta xmlns:x=\"adobe:ns:meta/\" xmlns:rdf=\"http://www.w3.org/1999/02/22-rdf-syntax-ns#\" " +
            "xmlns:iptc=\"http://iptc.org/std/Iptc4xmpExt/2008-02-29/\"><rdf:RDF><rdf:Description>" +
            "<x:Unrelated rdf:parseType=\"Literal\"><iptc:DigitalSourceType>trainedAlgorithmicMedia</iptc:DigitalSourceType></x:Unrelated>" +
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
}
