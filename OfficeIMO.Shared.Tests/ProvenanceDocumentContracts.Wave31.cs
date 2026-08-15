using OfficeIMO.Excel;
using OfficeIMO.Html;
using OfficeIMO.PowerPoint;
using OfficeIMO.Provenance;
using OfficeIMO.Visio;
using OfficeIMO.Word;
using Xunit;

namespace OfficeIMO.Shared.Tests;

public sealed partial class ProvenanceDocumentContracts {
    [Theory]
    [InlineData("word", "xlsx", "docx")]
    [InlineData("excel", "pptx", "xlsx")]
    [InlineData("powerpoint", "docx", "pptx")]
    [InlineData("visio", "docx", "vsdx")]
    public void OwningInspectionApisRejectOtherOfficePackageTypes(
        string target,
        string sourceExtension,
        string targetExtension) {
        byte[] package = CreateSavedOpenXmlPackage(sourceExtension);
        string path = Path.Combine(Path.GetTempPath(), Guid.NewGuid().ToString("N") + "." + targetExtension);
        try {
            File.WriteAllBytes(path, package);

            Assert.Throws<InvalidDataException>(() => target switch {
                "word" => WordDocument.InspectProvenance(path),
                "excel" => ExcelDocument.InspectProvenance(path),
                "powerpoint" => PowerPointPresentation.InspectProvenance(path),
                "visio" => VisioDocument.InspectProvenance(path),
                _ => throw new ArgumentOutOfRangeException(nameof(target))
            });
        } finally {
            if (File.Exists(path)) File.Delete(path);
        }
    }

    [Fact]
    public void HtmlProcessesSrcdocNativeManifestsWhenEmbeddedImagesAreDisabled() {
        string dataUri = "data:image/png;base64," + Convert.ToBase64String(CreatePngWithManifest(CreateManifestStore()));
        string nested = $"<html><head><script type=\"application/c2pa\">{Convert.ToBase64String(CreateManifestStore())}</script></head>" +
            $"<body><img src=\"{dataUri}\"></body></html>";
        string html = $"<html><body><iframe srcdoc=\"{System.Net.WebUtility.HtmlEncode(nested)}\"></iframe></body></html>";

        var inspectionOptions = new OfficeProvenanceOptions { ProcessEmbeddedAssets = false };
        OfficeProvenanceReport report = HtmlProvenance.Inspect(html, inspectionOptions);

        OfficeProvenanceEvidence evidence = Assert.Single(report.Evidence);
        Assert.Equal(OfficeProvenanceCarrierKind.C2paManifest, evidence.Carrier);
        Assert.StartsWith("HTML/iframe[srcdoc][0]", evidence.Location, StringComparison.Ordinal);

        var removalOptions = new OfficeProvenanceRemovalOptions { ProcessEmbeddedAssets = false };
        OfficeProvenanceRemovalResult result = HtmlProvenance.Remove(html, removalOptions);
        string output = System.Text.Encoding.UTF8.GetString(result.ToArray());

        Assert.Single(result.Before.Evidence);
        Assert.Empty(result.After.Evidence);
        Assert.Contains(dataUri, output, StringComparison.Ordinal);
    }
}
