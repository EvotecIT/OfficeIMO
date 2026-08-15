using System.Text;
using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.Provenance;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Fact]
    public void OpenXmlSignatureRelationshipsHonorTheXmlNodeBudgetBeforeDeletion() {
        string relationships = string.Concat(Enumerable.Range(0, 32).Select(index =>
            "<Relationship Id=\"rId" + index + "\" Type=\"http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/signature\" Target=\"sig1.xml\"/>"));
        byte[] package = ReplaceWave38Entry(
            CreateWave33XlsbProvenancePackage(signed: true),
            "_xmlsignatures/_rels/origin.sigs.rels",
            "<Relationships xmlns=\"http://schemas.openxmlformats.org/package/2006/relationships\">" +
            relationships + "</Relationships>");
        var options = new OfficeProvenanceRemovalOptions {
            SignatureMutationPolicy = OfficeIMO.OfficeSignatureMutationPolicy.RemoveInvalidatedSignatures
        };
        options.Limits.MaxContainerEntries = 16;

        Assert.Throws<InvalidDataException>(() =>
            ExcelDocument.RemoveProvenance(package, "workbook.xlsb", options));
    }

    [Fact]
    public void CssCommentsDoNotTurnCompoundCustomPropertySelectorsIntoDescendants() {
        string dataUri = "data:image/png;base64," +
            Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string html = "<html><head><style>.theme/**/.active { --hero: url('" + dataUri + "'); }" +
            ".target { background-image: var(--hero); }</style></head><body>" +
            "<div class=\"theme active\"><div class=\"target\"></div></div></body></html>";

        OfficeProvenanceReport report = HtmlProvenance.Inspect(html);
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html);

        Assert.True(report.HasC2paManifest);
        Assert.True(result.WasChanged);
        Assert.Empty(result.After.Evidence);
    }
}
