using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void EpubOwnershipRequiresABoundedOpfPackageDocument() {
        const string container =
            "<container xmlns=\"urn:oasis:names:tc:opendocument:xmlns:container\" version=\"1.0\"><rootfiles>" +
            "<rootfile full-path=\"OPS/package.opf\" media-type=\"application/oebps-package+xml\"/>" +
            "</rootfiles></container>";
        byte[] package = CreateEpubTestZip(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/container.xml", Encoding.UTF8.GetBytes(container)),
            ("OPS/package.opf", Encoding.UTF8.GetBytes("not an OPF package")),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        Assert.Throws<InvalidDataException>(() =>
            EpubDocument.RemoveProvenance(package, "publication.epub"));
    }

    [Fact]
    public void CssUrlsRetainNonCssUnicodeWhitespace() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><head><style>body{{background:url(\"\u00A0{dataUri}\")}}</style></head><body></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void SrcsetUsesHtmlAsciiWhitespaceForDescriptorBoundaries() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><body><img srcset=\"{dataUri}\u00A01x\"></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ExcelXlsbOwnershipRequiresTheExactWorkbookContentType() {
        byte[] package = ReplaceWave38Entry(
            CreateWave33XlsbProvenancePackage(signed: false),
            "[Content_Types].xml",
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.custom.binary\"/>" +
            "</Types>");

        Assert.ThrowsAny<Exception>(() =>
            ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));
    }
}
