using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void ExcelXlsbOwnershipRejectsForeignPackageMetadataNamespaces() {
        byte[] package = ReplaceWave38Entry(
            CreateWave33XlsbProvenancePackage(signed: false),
            "_rels/.rels",
            "<Relationships xmlns=\"urn:foreign:relationships\">" +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
            "</Relationships>");

        Assert.ThrowsAny<Exception>(() => ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));
    }

    [Fact]
    public void ExcelXlsbOwnershipRejectsForeignContentTypeNamespaces() {
        byte[] package = ReplaceWave38Entry(
            CreateWave33XlsbProvenancePackage(signed: false),
            "[Content_Types].xml",
            "<Types xmlns=\"urn:foreign:content-types\">" +
            "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>" +
            "</Types>");

        Assert.ThrowsAny<Exception>(() => ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));
    }

    [Fact]
    public void EpubOwnershipRequiresContainerMetadata() {
        byte[] package = CreateEpubTestZip(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        Assert.Throws<InvalidDataException>(() => EpubDocument.RemoveProvenance(package, "publication.epub"));
    }

    [Fact]
    public void EpubOwnershipRequiresAnExistingDeclaredRootfile() {
        const string container = "<container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" version=\"1.0\">" +
            "<rootfiles><rootfile full-path=\"OPS/missing.opf\" media-type=\"application/oebps-package+xml\"/></rootfiles></container>";
        byte[] package = CreateEpubTestZip(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/container.xml", Encoding.UTF8.GetBytes(container)),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        Assert.Throws<InvalidDataException>(() => EpubDocument.RemoveProvenance(package, "publication.epub"));
    }

    [Fact]
    public void HtmlRelationshipTokensUseOnlyAsciiWhitespace() {
        string html = "<html><head><link rel=\"stylesheet\u00A0c2pa-manifest\" href=\"claim.c2pa\"></head><body></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
        Assert.Equal(html, Encoding.UTF8.GetString(result.ToArray()));
    }

    [Fact]
    public void WordRelationshipResolvedApplicationSignatureMetadataBlocksRemoval() {
        byte[] package = CreateWordPackageWithNoncanonicalApplicationSignatureMetadata();

        InvalidOperationException exception = Assert.Throws<InvalidOperationException>(() =>
            WordDocument.RemoveProvenance(package, "document.docx"));

        Assert.Contains("invalidate package signatures", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    private static byte[] CreateWordPackageWithNoncanonicalApplicationSignatureMetadata() {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + ".docx");
        try {
            CreateOpenXmlPackage(path, "docx");
            using var output = new MemoryStream();
            byte[] original = File.ReadAllBytes(path);
            output.Write(original, 0, original.Length);
            output.Position = 0;
            using (Package package = Package.Open(output, FileMode.Open, FileAccess.ReadWrite)) {
                const string relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
                foreach (PackageRelationship relationship in package.GetRelationshipsByType(relationshipType).ToArray()) {
                    package.DeleteRelationship(relationship.Id);
                }
                Uri canonicalUri = PackUriHelper.CreatePartUri(new Uri("/docProps/app.xml", UriKind.Relative));
                if (package.PartExists(canonicalUri)) package.DeletePart(canonicalUri);
                Uri customUri = PackUriHelper.CreatePartUri(new Uri("/metadata/application.xml", UriKind.Relative));
                PackagePart application = package.CreatePart(
                    customUri,
                    "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                    CompressionOption.Maximum);
                using (var writer = new StreamWriter(application.GetStream(), new UTF8Encoding(false), 4096, leaveOpen: false)) {
                    writer.Write("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><DigSig>signature</DigSig></Properties>");
                }
                package.CreateRelationship(customUri, TargetMode.Internal, relationshipType);
                Uri manifestUri = PackUriHelper.CreatePartUri(new Uri("/META-INF/content_credential.c2pa", UriKind.Relative));
                using Stream target = package.CreatePart(manifestUri, "application/c2pa", CompressionOption.Maximum).GetStream();
                byte[] manifest = CreateManifestStore();
                target.Write(manifest, 0, manifest.Length);
            }
            return output.ToArray();
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }
}
