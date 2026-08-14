using System.IO.Compression;
using System.Text;
using OfficeIMO.Html;
using OfficeIMO.Markdown;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlProcessesEveryRepeatedBareStringImageSetSource() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><head><style>.x{{background:image-set(\"{dataUri}\" 1x, \"{dataUri}\" 2x)}}</style></head><body class=\"x\"></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Equal(2, report.Evidence.Count);
        Assert.Equal(2, result.Changes.Count);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void HtmlDomPreflightRequiresRawTextClosingTagDelimiter() {
        string falseClosings = string.Concat(Enumerable.Repeat("</scripture><div></div>", 16));
        string html = "<html><head><script>" + falseClosings + "</script></head><body><p>kept</p></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html,
            new OfficeProvenanceOptions { MaxContainerEntries = 8 });

        Assert.Empty(report.Evidence);
    }

    [Fact]
    public void MarkdownStringOverloadsRejectOversizedInputBeforeEncoding() {
        string markdown = new string('\u20ac', 128);
        var inspectOptions = new OfficeProvenanceOptions { MaxAssetBytes = 32, MaxManifestBytes = 16 };
        var removalOptions = new OfficeProvenanceRemovalOptions();
        removalOptions.Limits.MaxAssetBytes = 32;
        removalOptions.Limits.MaxManifestBytes = 16;

        Assert.Throws<InvalidDataException>(() => MarkdownProvenance.Inspect(markdown, inspectOptions));
        Assert.Throws<InvalidDataException>(() => MarkdownProvenance.Remove(markdown, removalOptions));
    }

    [Fact]
    public void OdfSignatureDetectionDoesNotSwitchToOpcForContentTypesEntry() {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, "META-INF/manifest.xml", ValidOdfManifestXml, CompressionLevel.Optimal);
                WriteEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>", CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/documentsignatures.xml", "<signatures/>", CompressionLevel.Optimal);
                WriteEntry(archive, "Pictures/provenance.png", CreatePngWithManifest(CreateManifestStore()), CompressionLevel.Optimal);
            }
            package = RewriteFixtureWithStoredMimetype(output.ToArray());
        }

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            OdfDocument.RemoveProvenance(package, "document.odt"));

        Assert.Contains("invalidate package signatures", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void HtmlDataUriFragmentsArePreservedOutsideThePayload() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore())) + "#icon";
        string html = $"<html><head></head><body><img src=\"{dataUri}\"></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("#icon", output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlManifestRemovalReportsPhysicalBytesAsUnknownAfterSerialization() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = $"<html><head><script type=\"application/c2pa\">{manifest}</script></head><body></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Equal(0, Assert.Single(result.Changes).RemovedBytes);
        Assert.True(result.WasReserialized);
    }
}
