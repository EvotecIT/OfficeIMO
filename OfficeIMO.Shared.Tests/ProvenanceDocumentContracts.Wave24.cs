using System.IO.Compression;
using System.Text;
using OfficeIMO;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlCustomPropertyRemovalPreservesOffsetsAcrossLeadingComments() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><head><style>/* legal */:root{{--hero:url({dataUri})}}.hero{{background:var(--hero)}}</style></head><body><div class=\"hero\"></div></body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
        Assert.Contains("/* legal */", output, StringComparison.Ordinal);
        Assert.DoesNotContain(dataUri, output, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlIgnoresCustomPropertyUsesInsideInactiveRules() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><head><style>:root{{--hero:url({dataUri})}}@media print{{.hero{{background:var(--hero)}}}}</style></head><body><div class=\"hero\"></div></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlManifestDecodingSharesTheExpandedBudgetAcrossSrcdocDocuments() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string nested = $"<html><head><script type=\"application/c2pa\">{manifest}</script></head></html>";
        string html = $"<html><head><script type=\"application/c2pa\">{manifest}</script></head><body><iframe srcdoc='{nested}'></iframe></body></html>";
        var options = new OfficeProvenanceOptions {
            MaxExpandedContainerBytes = CreateManifestStore().Length + 16
        };

        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Inspect(html, options));
    }

    [Fact]
    public void OdfSignatureManifestCleanupHonorsTheConfiguredXmlNodeBudget() {
        const string signaturePath = "META-INF/customsignatures.xml";
        string entries = string.Concat(Enumerable.Range(0, 32).Select(index =>
            $"<manifest:file-entry manifest:full-path=\"Pictures/{index}.png\" manifest:media-type=\"image/png\"/>"));
        string manifestXml =
            "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\">" +
            entries +
            $"<manifest:file-entry manifest:full-path=\"{signaturePath}\" manifest:media-type=\"text/xml\"/>" +
            "</manifest:manifest>";
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, signaturePath, "<signatures/>", CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/content_credential.c2pa", CreateManifestStore(), CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/manifest.xml", manifestXml, CompressionLevel.Optimal);
            }
            package = output.ToArray();
        }
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };
        options.Limits.MaxContainerEntries = 16;

        Assert.Throws<InvalidDataException>(() => OdfDocument.RemoveProvenance(package, "document.odt", options));
    }
}
