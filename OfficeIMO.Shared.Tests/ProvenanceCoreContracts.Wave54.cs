using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
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
