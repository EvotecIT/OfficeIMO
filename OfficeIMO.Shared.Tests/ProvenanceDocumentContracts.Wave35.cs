using System.IO.Compression;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void MarkdownOwnerDoesNotYieldToHtmlContentSniffing() {
        string markdown = "<!doctype html>\n\n-----BEGIN C2PA MANIFEST-----\n" +
            "data:application/c2pa;base64," + Convert.ToBase64String(CreateManifestStore()) + "\n" +
            "-----END C2PA MANIFEST-----\n";

        OfficeProvenanceReport report = MarkdownProvenance.Inspect(markdown);
        OfficeProvenanceRemovalResult result = MarkdownProvenance.Remove(markdown);

        Assert.Equal(OfficeProvenanceAssetFormat.StructuredText, report.Format);
        Assert.Single(report.Evidence);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("<!doctype html>", Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlPreflightTreatsQuotesInsideUnquotedValuesAsLiteral() {
        string html = "<html><body><div data-value=unquoted\">" +
            string.Concat(Enumerable.Repeat("<span></span>", 16)) +
            "</div></body></html>";

        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Inspect(
            html,
            new OfficeProvenanceOptions { MaxContainerEntries = 8 }));
    }

    [Fact]
    public void OdfRemovalValidatesPackageEntryLimitsBeforeManifestLookup() {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                for (int index = 0; index < 8; index++) WriteEntry(archive, $"Pictures/{index}.txt", "keep", CompressionLevel.Optimal);
            }
            package = RewriteFixtureWithStoredMimetype(output.ToArray());
        }
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxContainerEntries = 4;

        Assert.Throws<InvalidDataException>(() => OdfDocument.RemoveProvenance(package, "document.odt", options));
    }

    [Fact]
    public void VisioOwnershipValidationBoundsRootRelationshipXml() {
        byte[] package = CreateSignedVisioProvenancePackage(0);
        using var stream = new MemoryStream();
        stream.Write(package, 0, package.Length);
        stream.Position = 0;
        using (var archive = new ZipArchive(stream, ZipArchiveMode.Update, leaveOpen: true)) {
            archive.GetEntry("_rels/.rels")!.Delete();
            string relationships =
                "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
                string.Concat(Enumerable.Range(0, 24).Select(index =>
                    $"<Relationship Id=\"rId{index}\" Type=\"urn:test\" Target=\"part{index}.xml\"/>")) +
                "</Relationships>";
            WriteEntry(archive, "_rels/.rels", relationships, CompressionLevel.Optimal);
        }
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxContainerEntries = 16;

        Assert.Throws<InvalidDataException>(() => VisioDocument.RemoveProvenance(stream.ToArray(), "drawing.vsdx", options));
    }
}
