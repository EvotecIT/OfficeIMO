using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
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
