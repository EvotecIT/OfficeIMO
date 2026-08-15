using System.IO.Compression;
using OfficeIMO.Epub;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.OpenDocument;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Theory]
    [InlineData("word", "xlsx")]
    [InlineData("excel", "pptx")]
    [InlineData("powerpoint", "docx")]
    [InlineData("visio", "docx")]
    public void OwningRemovalApisRejectOtherOfficePackageTypes(string target, string sourceExtension) {
        byte[] package = CreateSavedOpenXmlPackage(sourceExtension);

        Assert.Throws<InvalidDataException>(() => target switch {
            "word" => WordDocument.RemoveProvenance(package, "document.docx"),
            "excel" => ExcelDocument.RemoveProvenance(package, "workbook.xlsx"),
            "powerpoint" => PowerPointPresentation.RemoveProvenance(package, "presentation.pptx"),
            "visio" => VisioDocument.RemoveProvenance(package, "drawing.vsdx"),
            _ => throw new ArgumentOutOfRangeException(nameof(target))
        });
    }

    [Theory]
    [InlineData("application/epub+zip", true)]
    [InlineData("application/vnd.oasis.opendocument.text", false)]
    public void PackageIdentityRejectsCompressedMimetypeEntries(string mimetype, bool epub) {
        byte[] package;
        using (var output = new MemoryStream()) {
            using (var archive = new ZipArchive(output, ZipArchiveMode.Create, leaveOpen: true)) {
                WriteEntry(archive, "mimetype", mimetype, CompressionLevel.Optimal);
            }
            package = output.ToArray();
        }

        Assert.Throws<InvalidDataException>(() => epub
            ? EpubDocument.RemoveProvenance(package, "publication.epub")
            : OdfDocument.RemoveProvenance(package, "document.odt"));
    }

    [Fact]
    public void UnsupportedHtmlImageDataUrisDoNotConsumeTheEmbeddedAssetBudget() {
        string unsupported = string.Concat(Enumerable.Repeat("<img src=\"data:image/avif;base64,AA==\">", 8));
        string supported = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><body>{unsupported}<img src=\"{supported}\"></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html,
            new OfficeProvenanceOptions { MaxEmbeddedAssets = 1 });

        Assert.Single(report.Evidence);
    }

    [Fact]
    public void HtmlPreflightTreatsForeignObjectChildrenAsHtml() {
        string html = "<html><body><svg><foreignObject><![CDATA[x>" +
            string.Concat(Enumerable.Repeat("<span></span>", 32)) +
            "</foreignObject></svg></body></html>";

        Assert.Throws<InvalidDataException>(() => HtmlProvenance.Inspect(
            html,
            new OfficeProvenanceOptions { MaxContainerEntries = 16 }));
    }

    private static byte[] CreateSavedOpenXmlPackage(string extension) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "." + extension);
        try {
            CreateOpenXmlPackage(path, extension);
            return File.ReadAllBytes(path);
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    private static byte[] RewriteFixtureWithStoredMimetype(byte[] package) {
        var entries = new List<OfficeProvenanceZipWriteEntry>();
        using (var archive = new ZipArchive(new MemoryStream(package, writable: false), ZipArchiveMode.Read)) {
            foreach (ZipArchiveEntry entry in archive.Entries) {
                using Stream input = entry.Open();
                using var content = new MemoryStream();
                input.CopyTo(content);
                byte[] data = content.ToArray();
                entries.Add(new OfficeProvenanceZipWriteEntry(
                    entry.FullName,
                    data.LongLength,
                    compress: !entry.FullName.Equals("mimetype", StringComparison.Ordinal),
                    entry.LastWriteTime,
                    internalAttributes: 0,
                    externalAttributes: unchecked((uint)entry.ExternalAttributes),
                    Array.Empty<byte>(),
                    Array.Empty<byte>(),
                    Array.Empty<byte>(),
                    () => new MemoryStream(data, writable: false)));
            }
        }
        return OfficeProvenanceZipWriter.Write(entries, long.MaxValue);
    }
}
