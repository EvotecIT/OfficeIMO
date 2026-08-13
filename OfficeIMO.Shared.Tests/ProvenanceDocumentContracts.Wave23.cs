using System.IO.Compression;
using System.Text;
using System.Xml.Linq;
using OfficeIMO;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlDoesNotUseCustomPropertyFromAnUnrelatedSelectorScope() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><head><style>.a{{--hero:url({dataUri})}}.b{{background:var(--hero)}}</style></head><body><div class=\"b\"></div></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlStringRemovalNormalizesLegacyCharsetDeclarationToUtf8() {
        string html = "<html><head><meta charset=\"windows-1252\"><script type=\"application/c2pa\">" +
            Convert.ToBase64String(CreateManifestStore()) + "</script></head><body>café</body></html>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string output = Encoding.UTF8.GetString(result.ToArray());

        Assert.Contains("charset=\"utf-8\"", output, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("windows-1252", output, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("café", output, StringComparison.Ordinal);
    }

    [Fact]
    public void OdfSignatureRemovalCleansManifestFileEntries() {
        const string signaturePath = "META-INF/customsignatures.xml";
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", "application/vnd.oasis.opendocument.text", CompressionLevel.NoCompression);
                WriteEntry(archive, signaturePath, "<signatures/>", CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/content_credential.c2pa", CreateManifestStore(), CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/manifest.xml",
                    "<manifest:manifest xmlns:manifest=\"urn:oasis:names:tc:opendocument:xmlns:manifest:1.0\"><manifest:file-entry manifest:full-path=\"/\" manifest:media-type=\"application/vnd.oasis.opendocument.text\"/><manifest:file-entry manifest:full-path=\"" + signaturePath + "\" manifest:media-type=\"text/xml\"/></manifest:manifest>",
                    CompressionLevel.Optimal);
            }
            package = output.ToArray();
        }

        OfficeProvenanceRemovalResult result = OdfDocument.RemoveProvenance(package, "document.odt", new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        });
        XDocument manifest = XDocument.Parse(Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "META-INF/manifest.xml")));

        Assert.DoesNotContain(manifest.Descendants(), element =>
            string.Equals((string?)element.Attribute(XName.Get("full-path", "urn:oasis:names:tc:opendocument:xmlns:manifest:1.0")), signaturePath, StringComparison.Ordinal));
    }

    [Fact]
    public void OpcRemovalCleansRelationshipsAndContentTypeForNativeManifest() {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "[Content_Types].xml", "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"><Override PartName=\"/META-INF/content_credential.c2pa\" ContentType=\"application/c2pa\"/></Types>", CompressionLevel.Optimal);
                WriteEntry(archive, "_rels/.rels", "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\"><Relationship Id=\"rId1\" Type=\"urn:c2pa\" Target=\"META-INF/content_credential.c2pa\"/></Relationships>", CompressionLevel.Optimal);
                WriteEntry(archive, "META-INF/content_credential.c2pa", CreateManifestStore(), CompressionLevel.Optimal);
            }
            package = output.ToArray();
        }

        OfficeProvenanceRemovalResult result = OfficeProvenanceRemover.Remove(package, "document.docx");

        Assert.DoesNotContain("content_credential.c2pa", Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "_rels/.rels")), StringComparison.Ordinal);
        Assert.DoesNotContain("content_credential.c2pa", Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "[Content_Types].xml")), StringComparison.Ordinal);
    }
}
