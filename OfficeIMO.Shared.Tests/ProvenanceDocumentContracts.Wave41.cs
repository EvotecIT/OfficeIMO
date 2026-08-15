using System.IO.Compression;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Theory]
    [InlineData("\u000B")]
    [InlineData("\u00A0")]
    public void HtmlUnquotedAttributesUseOnlyAsciiWhitespace(string nonHtmlWhitespace) {
        string html = "<svg><foreignObject x=a" + nonHtmlWhitespace + "/><![CDATA[" +
            string.Concat(Enumerable.Repeat("<div></div>", 32)) + "]]></foreignObject></svg>";

        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Inspect(
            html, new OfficeProvenanceOptions { MaxContainerEntries = 16 }));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("<root/>")]
    [InlineData("<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.spreadsheet\"/></manifest:manifest>")]
    public void OdfProvenanceRequiresAnOwningManifest(string? manifestXml) {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                if (manifestXml != null) WriteEntry(archive, "META-INF/manifest.xml", manifestXml, CompressionLevel.Optimal);
            }
            package = RewriteFixtureWithStoredMimetype(output.ToArray());
        }

        Assert.Throws<InvalidDataException>(() => OdfDocument.RemoveProvenance(package, "document.odt"));
    }
}
