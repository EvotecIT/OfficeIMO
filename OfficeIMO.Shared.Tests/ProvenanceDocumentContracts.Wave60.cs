using System.IO.Compression;
using System.IO.Packaging;
using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void AnimationShorthandKeywordsDoNotActivateSameNamedKeyframes() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>@keyframes linear{from{background-image:url('" + dataUri +
            "')}}.box{animation:1s linear other}</style><div class=\"box\"></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void AnimationShorthandCanUseAKeywordAsTheNameAfterItsComponentSlotIsFilled() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>@keyframes linear{from{background-image:url('" + dataUri +
            "')}}.box{animation:1s ease linear}</style><div class=\"box\"></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void LayeredCustomPropertyValuesRemainCaseSensitive() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string activeSvg = "<svg xmlns='http://www.w3.org/2000/svg' xmlns:c2pa='http://c2pa.org/manifest'>" +
            "<metadata><c2pa:manifest>" + manifest + "</c2pa:manifest></metadata></svg>";
        string inactiveSvg = activeSvg.Replace("c2pa", "C2PA").Replace("manifest", "MANIFEST");
        string active = "url('data:image/svg+xml," + Uri.EscapeDataString(activeSvg) + "')";
        string inactive = "url('data:image/svg+xml," + Uri.EscapeDataString(inactiveSvg) + "')";
        string html = "<style>.box{--hero:" + active + ";background-image:var(--hero)}" +
            "@layer low{.box{--hero:" + inactive + "}}</style><div class=\"box\"></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(result.Before.Evidence);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void PictureSourcesAfterTheFallbackImageAreInactive() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string html = "<picture><img src=\"fallback.png\"><source srcset=\"" + dataUri + "\"></picture>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(result.Before.Evidence);
        Assert.False(result.WasChanged);
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    public void OpenXmlAdaptersIgnoreOrphanConventionalApplicationMetadata(string extension) {
        byte[] package = CreateOpenXmlPackageWithOrphanApplicationSignatureMetadata(extension);

        OfficeProvenanceRemovalResult result = RemoveOpenXmlWithOptions(
            package, extension, new OfficeProvenanceRemovalOptions());

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void XlsbDetectionHonorsConfiguredMetadataByteLimits() {
        byte[] package = CreateWave33XlsbProvenancePackage(signed: false);
        string relationships =
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            new string(' ', 1024 * 1024) +
            "<Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.bin\"/>" +
            "</Relationships>";
        package = ReplaceWave38Entry(package, "_rels/.rels", relationships);

        OfficeProvenanceRemovalResult result = ExcelDocument.RemoveProvenance(package, "workbook.xlsb");

        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }

    private static byte[] CreateOpenXmlPackageWithOrphanApplicationSignatureMetadata(string extension) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "." + extension);
        try {
            CreateOpenXmlPackage(path, extension);
            byte[] original = File.ReadAllBytes(path);
            using var output = new MemoryStream();
            output.Write(original, 0, original.Length);
            output.Position = 0;
            using (Package package = Package.Open(output, FileMode.Open, FileAccess.ReadWrite)) {
                const string relationshipType = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties";
                foreach (PackageRelationship relationship in package.GetRelationshipsByType(relationshipType).ToArray()) {
                    package.DeleteRelationship(relationship.Id);
                }
                Uri applicationUri = PackUriHelper.CreatePartUri(new Uri("/docProps/app.xml", UriKind.Relative));
                PackagePart application = package.PartExists(applicationUri)
                    ? package.GetPart(applicationUri)
                    : package.CreatePart(
                        applicationUri,
                        "application/vnd.openxmlformats-officedocument.extended-properties+xml",
                        CompressionOption.Maximum);
                using (var writer = new StreamWriter(application.GetStream(FileMode.Create, FileAccess.Write), new UTF8Encoding(false), 4096, leaveOpen: false)) {
                    writer.Write("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\"><DigSig>signature</DigSig></Properties>");
                }
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
