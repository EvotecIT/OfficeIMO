using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void SvgDoesNotCountOneXmpRootTwiceWhenItHasOtherIptcMetadata() {
        byte[] svg = Encoding.UTF8.GetBytes(
            "<svg xmlns='http://www.w3.org/2000/svg' xmlns:x='adobe:ns:meta/' " +
            "xmlns:rdf='http://www.w3.org/1999/02/22-rdf-syntax-ns#' " +
            "xmlns:iptc='http://iptc.org/std/Iptc4xmpExt/2008-02-29/'>" +
            "<metadata><x:xmpmeta iptc:Note='other'><rdf:RDF><rdf:Description " +
            "iptc:DigitalSourceType='trainedAlgorithmicMedia'/></rdf:RDF></x:xmpmeta></metadata></svg>");

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(svg, "fixture.svg");

        Assert.Single(result.Before.Evidence);
        Assert.True(result.WasChanged);
        Assert.False(result.After.HasGenerativeAiDeclaration);
        Assert.Contains("iptc:Note", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void ZipRewriteKeepsAColonInsideASafeRelativeSegment() {
        byte[] package = CreateZipWithUnicodePathEntry(
            "Pictures/legacy.png", "Pictures/schema:v1.png", Encoding.UTF8.GetBytes("keep"));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "publication.zip");

        Assert.True(result.WasChanged);
        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("Pictures/schema:v1.png"));
    }
}
