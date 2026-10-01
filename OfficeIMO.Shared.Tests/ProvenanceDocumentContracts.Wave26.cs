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
    [Theory]
    [MemberData(nameof(HtmlPreflightKeepsRawAndForeignMarkupLiteralWithNativeManifestCases))]
    public void HtmlPreflightKeepsRawAndForeignMarkupLiteralWithNativeManifest(string caseName, string html, int maximumEntries) {
        _ = caseName;
        OfficeProvenanceReport report = HtmlProvenance.Inspect(
            html, new OfficeProvenanceOptions { MaxContainerEntries = maximumEntries });

        Assert.Single(report.Evidence);
    }

    public static IEnumerable<object[]> HtmlPreflightKeepsRawAndForeignMarkupLiteralWithNativeManifestCases() {
        {
            string manifest = Convert.ToBase64String(CreateManifestStore());
            string html = "<html><head><script type=\"application/c2pa\">" + manifest +
                "</script></head><body><plaintext>" + string.Concat(Enumerable.Repeat("<div>literal</div>", 128));
            yield return new object[] { "HtmlPreflightTreatsPlaintextRemainderAsText", html, 32 };
        }
        {
            string manifest = Convert.ToBase64String(CreateManifestStore());
            string html = "<html><head><script type=\"application/c2pa\">" + manifest +
                "</script></head><body><xmp>" + string.Concat(Enumerable.Repeat("<div>literal</div>", 128)) +
                "</xmp></body></html>";
            yield return new object[] { "HtmlPreflightTreatsLegacyRawTextElementsAsText", html, 32 };
        }
        {
            string manifest = Convert.ToBase64String(CreateManifestStore());
            string html = "<html><head><script type=\"application/c2pa\">" + manifest +
                "</script></head><body><svg><![CDATA[" + string.Concat(Enumerable.Repeat("<div></div>", 64)) +
                "]]></svg></body></html>";
            yield return new object[] { "HtmlPreflightPreservesForeignContentCdataAsText", html, 16 };
        }
        {
            string manifest = Convert.ToBase64String(CreateManifestStore());
            string html = "<html><head><script type=\"application/c2pa\">" + manifest +
                "</script></head><body><math><mi><mglyph><![CDATA[" +
                string.Concat(Enumerable.Repeat("<div></div>", 64)) +
                "]]></mglyph></mi></math></body></html>";
            yield return new object[] { "HtmlPreflightKeepsMathMlGlyphCdataInForeignContent", html, 32 };
        }
        {
            string manifest = Convert.ToBase64String(CreateManifestStore());
            string html = "<html><head><script type=\"application/c2pa\">" + manifest +
                "</script></head><body><script><!--<script></script>" +
                string.Concat(Enumerable.Repeat("<div></div>", 64)) +
                "</script></body></html>";
            yield return new object[] { "HtmlPreflightModelsScriptDoubleEscapedState", html, 32 };
        }
        {
            string manifest = Convert.ToBase64String(CreateManifestStore());
            string html = "<html><head><script type=\"application/c2pa\">" + manifest +
                "</script></head><body><svg><font title=\" color=x\"><![CDATA[" +
                string.Concat(Enumerable.Repeat("<div></div>", 64)) +
                "]]></font></svg></body></html>";
            yield return new object[] { "HtmlForeignFontBreakoutUsesAttributeNamesNotQuotedValues", html, 32 };
        }
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
