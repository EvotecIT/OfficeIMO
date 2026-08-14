using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void HtmlPreflightDecodesMathMlIntegrationPointEncoding() {
        string html =
            "<math><annotation-xml encoding=\"text&#x2f;html\"><![CDATA[>" +
            "<span></span><span></span>]]></annotation-xml></math>";

        AssertHtmlPreflightRejects(html, maximumEntries: 3);
    }

    [Fact]
    public void DataUriBase64MarkerMustBeTheFinalMetadataSegment() {
        string payload = Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><body><img src=\"data:image/png;base64;charset=utf-8,{payload}\"></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.Empty(report.Evidence);
        Assert.False(result.WasChanged);
    }

    [Fact]
    public void ExcelXlsbSignatureCleanupPreservesUnrelatedTypedParts() {
        byte[] package = ReplaceWave38Entry(
            CreateWave33XlsbProvenancePackage(signed: true),
            "[Content_Types].xml",
            "<Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">" +
            "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>" +
            "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-excel.sheet.binary.macroEnabled.main\"/>" +
            "<Override PartName=\"/docProps/app.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.extended-properties+xml\"/>" +
            "<Override PartName=\"/_xmlsignatures/origin.sigs\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-origin\"/>" +
            "<Override PartName=\"/_xmlsignatures/sig1.xml\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "<Override PartName=\"/custom/unrelated.bin\" ContentType=\"application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml\"/>" +
            "</Types>");
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeIMO.OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };

        OfficeProvenanceRemovalResult result = ExcelDocument.RemoveProvenance(package, "workbook.xlsb", options);
        string contentTypes = Encoding.UTF8.GetString(ReadZipEntry(result.ToArray(), "[Content_Types].xml"));

        Assert.True(result.WereInvalidatedSignaturesRemoved);
        Assert.Contains("/custom/unrelated.bin", contentTypes, StringComparison.Ordinal);
    }

    [Fact]
    public void InvalidSrcsetDescriptorCandidateIsInert() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = $"<html><body><img srcset=\"{dataUri} 1x 2x\"></body></html>";

        Assert.Empty(HtmlSrcSetParser.Parse(dataUri + " 1x 2x"));
        Assert.Empty(HtmlProvenance.Inspect(html).Evidence);
        Assert.False(HtmlProvenance.Remove(html).WasChanged);
    }
}
