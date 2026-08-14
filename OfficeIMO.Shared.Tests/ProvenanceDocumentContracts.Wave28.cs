using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using OfficeIMO;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void OdfApiRejectsForeignOpcPackageBeforeMutation() {
        byte[] package = CreateEpubTestZip(
            ("[Content_Types].xml", Encoding.UTF8.GetBytes("<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>")),
            ("_xmlsignatures/sig1.xml", Encoding.UTF8.GetBytes("<Signature/>")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        Assert.Throws<InvalidDataException>(() => OdfDocument.RemoveProvenance(package, "document.odt"));
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".odt");
        try {
            File.WriteAllBytes(path, package);
            Assert.Throws<InvalidDataException>(() => OdfDocument.InspectProvenance(path));
        } finally {
            File.Delete(path);
        }
    }

    [Fact]
    public void HtmlPreflightTreatsLegacyRawTextElementsAsText() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<html><head><script type=\"application/c2pa\">" + manifest +
            "</script></head><body><xmp>" + string.Concat(Enumerable.Repeat("<div>literal</div>", 128)) +
            "</xmp></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html, new OfficeProvenanceOptions { MaxContainerEntries = 12 });

        Assert.Single(report.Evidence);
    }

    [Fact]
    public void HtmlIgnoresInactiveImagePreloads() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><head><link rel=\"preload\" as=\"image\" media=\"print\" href=\"{dataUri}\"></head><body></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    public void OpenXmlApplicationSignatureMetadataBlocksDefaultRemoval(string extension) {
        byte[] package = CreateOpenXmlPackageWithLargeApplicationMetadata(extension, 0);

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            RemoveOpenXmlWithOptions(package, extension, new OfficeProvenanceRemovalOptions()));

        Assert.Contains("invalidate package signatures", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void VisioApplicationSignatureMetadataBlocksDefaultRemoval() {
        byte[] package = CreateVisioPackageWithApplicationSignatureOnly();

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            VisioDocument.RemoveProvenance(package, "drawing.vsdx"));

        Assert.Contains("invalidate package signatures", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] CreateVisioPackageWithApplicationSignatureOnly(bool relationshipOwned = true) {
        using var output = new MemoryStream();
        using (Package package = Package.Open(output, FileMode.Create, FileAccess.ReadWrite)) {
            Uri documentUri = PackUriHelper.CreatePartUri(new Uri("/visio/document.xml", UriKind.Relative));
            using (Stream document = package.CreatePart(documentUri, "application/vnd.ms-visio.drawing.main+xml", CompressionOption.Maximum).GetStream()) {
                byte[] xml = Encoding.UTF8.GetBytes("<VisioDocument xmlns=\"http://schemas.microsoft.com/office/visio/2012/main\"/>");
                document.Write(xml, 0, xml.Length);
            }
            package.CreateRelationship(documentUri, TargetMode.Internal, "http://schemas.microsoft.com/visio/2010/relationships/document");
            Uri manifestUri = PackUriHelper.CreatePartUri(new Uri("/META-INF/content_credential.c2pa", UriKind.Relative));
            using (Stream target = package.CreatePart(manifestUri, "application/c2pa", CompressionOption.Maximum).GetStream()) {
                byte[] manifest = CreateManifestStore();
                target.Write(manifest, 0, manifest.Length);
            }
            Uri appUri = PackUriHelper.CreatePartUri(new Uri("/docProps/app.xml", UriKind.Relative));
            PackagePart app = package.CreatePart(
                appUri,
                "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                CompressionOption.Maximum);
            if (relationshipOwned) {
                package.CreateRelationship(
                    appUri,
                    TargetMode.Internal,
                    "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties");
            }
            using var writer = new StreamWriter(app.GetStream(), new UTF8Encoding(false), 4096, leaveOpen: false);
            writer.Write("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><DigSig>signature</DigSig></Properties>");
        }
        return output.ToArray();
    }
}
