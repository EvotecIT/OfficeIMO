using System.IO.Compression;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void OdfProvenanceRequiresTheCanonicalContentPart() {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, "META-INF/manifest.xml", ValidOdfManifestXml, CompressionLevel.Optimal);
            }
            package = RewriteFixtureWithStoredMimetype(output.ToArray());
        }

        Assert.Throws<InvalidDataException>(() => OdfDocument.RemoveProvenance(package, "document.odt"));
    }

    [Fact]
    public void HtmlClosingTagAttributesHonorQuotedGreaterThanCharacters() {
        string quotedMarkup = string.Concat(Enumerable.Repeat("<span>", 16));
        string html = "<div></div x=\">" + quotedMarkup + "\">";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html,
            new OfficeProvenanceOptions { MaxContainerEntries = 4 });

        Assert.Empty(report.Evidence);
    }
}
