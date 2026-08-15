using System.IO.Compression;
using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void ResponsivePictureSourcesRemainInProvenanceScope() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<picture><source media='(max-width:600px)' srcset='" + dataUri +
            "'><img src='fallback.png'></picture>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void EpubProvenanceRejectsUnsafeArchiveEntryPaths() {
        string container = Wave63ContainerPrefix +
            "<rootfile full-path='package.opf' media-type='application/oebps-package+xml'/>" +
            Wave63ContainerSuffix;
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/container.xml", Encoding.UTF8.GetBytes(container)),
            ("package.opf", Encoding.UTF8.GetBytes(Wave63Opf)),
            ("Pictures/../image.png", CreatePngWithManifest(CreateManifestStore())));

        Assert.Throws<InvalidDataException>(() => EpubDocument.RemoveProvenance(package));
    }

    [Fact]
    public void OdfPreservesManifestDeclarationForRetainedMalformedNativeCarrier() {
        const string odfManifest =
            "<manifest:manifest xmlns:manifest='urn:oasis:names:tc:opendocument:xmlns:manifest:1.0'>" +
            "<manifest:file-entry manifest:full-path='/' manifest:media-type='application/vnd.oasis.opendocument.text'/>" +
            "<manifest:file-entry manifest:full-path='META-INF/content_credential.c2pa' manifest:media-type='application/c2pa'/>" +
            "</manifest:manifest>";
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes("application/vnd.oasis.opendocument.text")),
            ("META-INF/manifest.xml", Encoding.UTF8.GetBytes(odfManifest)),
            ("content.xml", Encoding.UTF8.GetBytes("<content/>")),
            ("META-INF/content_credential.c2pa", new byte[] { 1, 2, 3 }),
            ("Pictures/provenance.png", CreatePngWithManifest(CreateManifestStore())));

        OfficeProvenanceRemovalResult result = OdfDocument.RemoveProvenance(package, "document.odt");

        using var archive = new ZipArchive(new MemoryStream(result.ToArray()), ZipArchiveMode.Read);
        Assert.NotNull(archive.GetEntry("META-INF/content_credential.c2pa"));
        Assert.Contains("content_credential.c2pa", Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "META-INF/manifest.xml")), StringComparison.Ordinal);
    }

    [Fact]
    public void PackageByteRemovalEnforcesTheOuterAssetLimit() {
        byte[] package = CreateZipPackage(
            "odt",
            "META-INF/documentsignatures.xml",
            CreatePngWithManifest(CreateManifestStore()));
        var options = new OfficeProvenanceRemovalOptions();
        options.Limits.MaxAssetBytes = package.LongLength - 1;
        options.Limits.MaxManifestBytes = options.Limits.MaxAssetBytes;

        InvalidDataException exception = Assert.Throws<InvalidDataException>(() =>
            OdfDocument.RemoveProvenance(package, "document.odt", options));

        Assert.Contains("package exceeds", exception.Message, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void DataUriStylesheetFragmentsRemainUrlFragmentsAfterRewrite() {
        string image = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string css = ".box{background-image:url('" + image + "')}";
        string stylesheet = "data:text/css;base64," +
            Convert.ToBase64String(Encoding.UTF8.GetBytes(css)) + "#theme";
        string html = "<link rel='stylesheet' href='" + stylesheet + "'><div class='box'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);
        string rewritten = Encoding.UTF8.GetString(result.ToArray());

        Assert.True(result.WasChanged);
        Assert.Contains("#theme", rewritten, StringComparison.Ordinal);
    }
}
