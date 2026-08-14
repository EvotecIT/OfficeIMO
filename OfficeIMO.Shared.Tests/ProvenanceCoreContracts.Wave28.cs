using System.IO.Compression;
using System.Text;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceCoreContracts {
    [Fact]
    public void PngStrictRemovalPreservesDuplicateC2paChunks() {
        byte[] png = Join(
            new byte[] { 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A },
            CreatePngChunk("IHDR", new byte[13]),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("caBX", CreateManifestStore()),
            CreatePngChunk("IEND", Array.Empty<byte>()));

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(png, "fixture.png");

        Assert.Equal(2, result.Before.Evidence.Count);
        Assert.All(result.Before.Evidence, item => Assert.False(item.IsStructurallyValid));
        Assert.False(result.WasChanged);
        Assert.Equal(png, result.ToArray());
    }

    [Fact]
    public void CompetingTextCarriersReportEvidenceInSourceOrderAndArePreserved() {
        byte[] wrapper = CreateTextWrapper(CreateManifestStore());
        byte[] structured = Encoding.UTF8.GetBytes(
            "\n-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n");
        byte[] text = Join(wrapper, structured);

        OfficeProvenanceReport report = OfficeProvenanceInspector.Inspect(text, "fixture.txt");
        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(text, "fixture.txt");

        Assert.Equal(2, report.Evidence.Count);
        Assert.Contains("C2PATextManifestWrapper", report.Evidence[0].Location, StringComparison.Ordinal);
        Assert.Contains("Text/C2PA@", report.Evidence[1].Location, StringComparison.Ordinal);
        Assert.All(report.Evidence, evidence => Assert.False(evidence.IsStructurallyValid));
        Assert.Empty(result.Changes);
        Assert.False(result.WasChanged);
        Assert.Equal(text, result.ToArray());
    }

    [Fact]
    public void ZipSignatureRelationshipInspectionSharesTheExpansionBudget() {
        byte[] relationship = Encoding.UTF8.GetBytes(
            "<Relationships><Relationship Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin\" Target=\"origin.sigs\"/>" +
            new string(' ', 4096) + "</Relationships>");
        byte[] package = CreateCompressedZip(
            ("META-INF/content_credential.c2pa", CreateManifestStore()),
            ("_rels/.rels", relationship));
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxAssetBytes = 16 * 1024;
        options.Limits.MaxManifestBytes = 1024;
        options.Limits.MaxExpandedContainerBytes = 512;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OfficeProvenanceRemover.Remove(package, "fixture.zip", options));

        Assert.Contains("expanded", exception.Message, StringComparison.OrdinalIgnoreCase);
    }
}
