using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRenderCapabilityGallery_BindsInputPdfPreviewsDiagnosticsAndExpectations() {
        const string html = """
            <style>
            @page { size: 5in 4in; margin: 24px; @top-center { content: "Quarterly report" } }
            body { margin: 0; font: 11px/15px Arial; }
            .columns { columns: 2; column-gap: 14px; column-rule: 1px solid #789; }
            table { border-collapse: collapse; width: 100%; } th,td { border: 1px solid #789; padding: 3px; }
            .vertical { writing-mode: vertical-rl; height: 72px; color: #174a7e; }
            .note { float: footnote; font-size: 8px; }
            </style>
            <h1>Quarterly report</h1>
            <div class="columns"><p>Column one content with enough words to wrap naturally.</p><p>Column two content.</p></div>
            <table><thead><tr><th>Item</th><th>Total</th></tr></thead><tbody><tr><td>Services</td><td>42</td></tr></tbody></table>
            <p class="vertical">縦書き PDF</p><p>Evidence<span class="note">Managed paged footnote.</span></p>
            """;
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Html.RenderGallery." + Guid.NewGuid().ToString("N"));
        try {
            var galleryOptions = new HtmlRenderCapabilityGalleryOptions(new HtmlCapabilityGalleryScenario(
                "managed-report",
                "Managed report",
                "HTML PDF",
                "Cross-artifact proof for paged layout, tables, columns, vertical text, and footnotes.")) {
                RenderOptions = new HtmlPdfSaveOptions {
                    PageSize = new OfficePageSize(5D, 4D),
                    Margins = HtmlRenderMargins.All(24D),
                    HonorCssPageRules = true
                }
            };
            foreach (string capability in new[] { "layout-columns", "layout-tables", "vertical-text", "paged-footnotes", "paged-page-rules" }) {
                galleryOptions.Expectations.Add(new HtmlCapabilityGalleryExpectation(
                    capability,
                    HtmlCapabilityGalleryExpectationOutcome.VisualProof,
                    "Hash-bound PDF, PNG, and SVG artifacts plus searchable PDF text."));
            }

            HtmlCapabilityGalleryManifest manifest = HtmlConversionDocument.Parse(html).SaveRenderCapabilityGallery(directory, galleryOptions);
            string prefix = Path.Combine(directory, "managed-report");
            byte[] pdf = File.ReadAllBytes(prefix + ".pdf");
            string extracted = PdfCore.PdfReadDocument.Open(pdf).ExtractText();
            using JsonDocument json = JsonDocument.Parse(File.ReadAllText(prefix + ".manifest.json"));

            Assert.Equal(4, manifest.Result.Artifacts.Count);
            Assert.All(manifest.Result.Artifacts, artifact => Assert.Equal(64, artifact.Sha256.Length));
            Assert.Contains("Quarterly report", extracted, StringComparison.Ordinal);
            Assert.Contains("Managed paged footnote", extracted, StringComparison.Ordinal);
            Assert.Equal("officeimo.html.capability-gallery", json.RootElement.GetProperty("schemaId").GetString());
            Assert.Equal(5, json.RootElement.GetProperty("expectations").GetArrayLength());
            Assert.True(File.Exists(prefix + ".preview.png"));
            Assert.True(File.Exists(prefix + ".preview.svg"));
            Assert.True(File.Exists(prefix + ".manifest.md"));
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }
}
