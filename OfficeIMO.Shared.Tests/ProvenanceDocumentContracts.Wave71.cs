using System.IO.Compression;
using System.Text;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Theory]
    [InlineData("_rels/.rels", "_rels\\.rels")]
    [InlineData("xl/workbook.bin", "xl\\workbook.bin")]
    public void ExcelXlsbRejectsBackslashZipEntryNames(string originalName, string replacementName) {
        byte[] package = RenameWave71Entry(
            CreateWave33XlsbProvenancePackage(signed: false),
            originalName,
            replacementName);

        Assert.ThrowsAny<Exception>(() => ExcelDocument.RemoveProvenance(package, "workbook.xlsb"));
    }

    [Fact]
    public void HtmlSanitizesActiveShapeOutsideImageCarriers() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<style>.cutout{shape-outside:url('" + dataUri + "')}</style><div class='cutout'></div>";

        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(result.WasChanged);
        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
    }

    [Fact]
    public void EpubRejectsAbsoluteIriRootfilePaths() {
        const string path = "http:package.opf";
        string container = Wave63ContainerPrefix +
            "<rootfile full-path='" + path + "' media-type='application/oebps-package+xml'/>" +
            Wave63ContainerSuffix;
        byte[] package = CreateStoredPackage(
            ("mimetype", Encoding.ASCII.GetBytes("application/epub+zip")),
            ("META-INF/container.xml", Encoding.UTF8.GetBytes(container)),
            (path, Encoding.UTF8.GetBytes(Wave63Opf)),
            ("META-INF/content_credential.c2pa", CreateManifestStore()));

        Assert.Throws<InvalidDataException>(() => EpubDocument.RemoveProvenance(package));
    }

    private static byte[] RenameWave71Entry(byte[] package, string originalName, string replacementName) {
        using var source = new ZipArchive(new MemoryStream(package), ZipArchiveMode.Read);
        using var output = new MemoryStream();
        using (var target = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
            foreach (ZipArchiveEntry sourceEntry in source.Entries) {
                string name = string.Equals(sourceEntry.FullName, originalName, StringComparison.Ordinal)
                    ? replacementName
                    : sourceEntry.FullName;
                ZipArchiveEntry targetEntry = target.CreateEntry(name, CompressionLevel.Optimal);
                using Stream input = sourceEntry.Open();
                using Stream destination = targetEntry.Open();
                input.CopyTo(destination);
            }
        }
        return output.ToArray();
    }
}
