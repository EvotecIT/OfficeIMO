using System.IO.Compression;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlPreflightTreatsPlaintextRemainderAsText() {
        string manifest = Convert.ToBase64String(CreateManifestStore());
        string html = "<html><head><script type=\"application/c2pa\">" + manifest +
            "</script></head><body><plaintext>" + string.Concat(Enumerable.Repeat("<div>literal</div>", 128));

        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html, new OfficeProvenanceOptions { MaxContainerEntries = 32 });

        Assert.Single(report.Evidence);
    }

    [Fact]
    public void HtmlSkipsResolvedCustomPropertyUrlFallbacks() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string html = "<html><head><style>:root{--hero:none}.hero{background-image:var(--hero,url(" +
            dataUri + "))}</style></head><body><div class=\"hero\"></div></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void HtmlRestrictsLegacyBackgroundImagesToSupportedElements() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(
            CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><body><div background=\"{dataUri}\"></div><table background=\"{dataUri}\"><tr><td>x</td></tr></table></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Single(report.Evidence);
        Assert.True(result.WasChanged);
        Assert.Contains(dataUri, Encoding.UTF8.GetString(result.ToArray()), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("docx")]
    [InlineData("xlsx")]
    [InlineData("pptx")]
    public void OpenXmlSignatureCleanupBoundsApplicationMetadata(string extension) {
        byte[] package = CreateOpenXmlPackageWithLargeApplicationMetadata(extension, 300);
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };
        options.Limits.MaxContainerEntries = 256;

        Assert.Throws<InvalidDataException>(() => RemoveOpenXmlWithOptions(package, extension, options));
    }

    private static byte[] CreateOpenXmlPackageWithLargeApplicationMetadata(string extension, int elementCount) {
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "." + extension);
        try {
            CreateOpenXmlPackage(path, extension);
            var xml = new StringBuilder("<Properties xmlns=\"http://schemas.openxmlformats.org/officeDocument/2006/extended-properties\">");
            for (int index = 0; index < elementCount; index++) xml.Append("<Item Index=\"").Append(index).Append("\"/>");
            xml.Append("<DigSig>signature</DigSig></Properties>");
            using (FileStream packageStream = File.Open(path, FileMode.Open, FileAccess.ReadWrite)) {
                WriteApplicationProperties(packageStream, extension, Encoding.UTF8.GetBytes(xml.ToString()));
            }
            using (ZipArchive archive = ZipFile.Open(path, ZipArchiveMode.Update)) {
                WriteEntry(archive, "META-INF/content_credential.c2pa", CreateManifestStore(), CompressionLevel.Optimal);
            }
            return File.ReadAllBytes(path);
        } finally {
            File.Delete(path);
        }
    }

    private static void WriteApplicationProperties(Stream package, string extension, byte[] xml) {
        ExtendedFilePropertiesPart part;
        switch (extension) {
            case "docx":
                using (WordprocessingDocument document = WordprocessingDocument.Open(package, true)) {
                    part = document.ExtendedFilePropertiesPart ?? document.AddExtendedFilePropertiesPart();
                    WritePart(part, xml);
                }
                break;
            case "xlsx":
                using (SpreadsheetDocument document = SpreadsheetDocument.Open(package, true)) {
                    part = document.ExtendedFilePropertiesPart ?? document.AddExtendedFilePropertiesPart();
                    WritePart(part, xml);
                }
                break;
            case "pptx":
                using (PresentationDocument document = PresentationDocument.Open(package, true)) {
                    part = document.ExtendedFilePropertiesPart ?? document.AddExtendedFilePropertiesPart();
                    WritePart(part, xml);
                }
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(extension));
        }
    }

    private static void WritePart(OpenXmlPart part, byte[] xml) {
        using Stream output = part.GetStream(FileMode.Create, FileAccess.Write);
        output.Write(xml, 0, xml.Length);
    }

    private static OfficeProvenanceRemovalResult RemoveOpenXmlWithOptions(
        byte[] package,
        string extension,
        OfficeProvenanceRemovalOptions options) => extension switch {
            "docx" => WordDocument.RemoveProvenance(package, "document.docx", options),
            "xlsx" => ExcelDocument.RemoveProvenance(package, "workbook.xlsx", options),
            "pptx" => PowerPointPresentation.RemoveProvenance(package, "presentation.pptx", options),
            _ => throw new ArgumentOutOfRangeException(nameof(extension))
        };
}
