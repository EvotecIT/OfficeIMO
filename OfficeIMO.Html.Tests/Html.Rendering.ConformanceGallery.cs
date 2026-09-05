using System.Text.Json;
using OfficeIMO.Drawing;
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed partial class HtmlRenderingTests {
    [Fact]
    public void HtmlRenderCapabilityGallery_RecordsAllPagesFormatsAndFailedExecutedProof() {
        const string html = "<style>@page { size: 200px 150px; margin: 10px; } body { margin: 0; }</style>" +
            "<p>First page</p><p style='break-before:page;transform:rotateX(30deg)'>Second page</p>";
        string directory = Path.Combine(Path.GetTempPath(), "OfficeIMO.Html.AllPageGallery." + Guid.NewGuid().ToString("N"));
        try {
            var options = new HtmlRenderCapabilityGalleryOptions(new HtmlCapabilityGalleryScenario(
                "all-pages", "All pages", "Rendering", "Executed artifact checks")) {
                PreviewAllPages = true
            };
            options.PreviewFormats.Clear();
            foreach (OfficeImageExportFormat format in new[] { OfficeImageExportFormat.Png, OfficeImageExportFormat.Jpeg,
                OfficeImageExportFormat.Tiff, OfficeImageExportFormat.Webp, OfficeImageExportFormat.Svg }) options.PreviewFormats.Add(format);
            options.PdfProofOptions.RequiredPageCount = 2;
            options.PdfProofOptions.RequiredTextMarkers.Add("Deliberately absent marker");
            options.Expectations.Add(new HtmlCapabilityGalleryExpectation("unverified-visual", HtmlCapabilityGalleryExpectationOutcome.VisualProof, "Manual review required"));

            HtmlCapabilityGalleryManifest manifest = HtmlConversionDocument.Parse(html).SaveRenderCapabilityGallery(directory, options);
            Assert.Equal(12, manifest.Result.Artifacts.Count);
            HtmlCapabilityGalleryArtifact pdf = Assert.Single(manifest.Result.Artifacts, artifact => artifact.Id == "pdf");
            Assert.Equal(2, pdf.Evidence!.PageCount);
            Assert.False(Assert.Single(pdf.Evidence.Checks).Passed);
            Assert.Contains("Deliberately absent marker", Assert.Single(pdf.Evidence.Checks).Detail, StringComparison.Ordinal);
            HtmlCapabilityGalleryArtifact[] images = manifest.Result.Artifacts.Where(artifact => artifact.Evidence?.PageNumber != null).ToArray();
            Assert.Equal(10, images.Length);
            Assert.All(images, artifact => {
                OfficeImageInfo identified = OfficeImageReader.Identify(File.ReadAllBytes(artifact.Path));
                Assert.Equal(identified.Width, artifact.Evidence!.Width);
                Assert.Equal(identified.Height, artifact.Evidence.Height);
                Assert.Equal(2, artifact.Evidence.PageCount);
                Assert.True(Assert.Single(artifact.Evidence.Checks).Passed);
                Assert.True(artifact.Evidence.HasLoss);
                Assert.Contains(artifact.Evidence.Diagnostics, diagnostic => diagnostic.LossKind == OfficeConversionLossKind.Omission);
            });
            Assert.Equal(5, images.Count(artifact => artifact.Evidence!.PageNumber == 1));
            Assert.Equal(5, images.Count(artifact => artifact.Evidence!.PageNumber == 2));
            using JsonDocument json = JsonDocument.Parse(File.ReadAllText(Path.Combine(directory, "all-pages.manifest.json")));
            Assert.Equal("1.1", json.RootElement.GetProperty("schemaVersion").GetString());
            Assert.Equal("declared-not-executed", json.RootElement.GetProperty("expectationStatus").GetString());
            Assert.False(json.RootElement.GetProperty("artifacts")[1].GetProperty("evidence").GetProperty("checks")[0].GetProperty("passed").GetBoolean());
        } finally {
            if (Directory.Exists(directory)) Directory.Delete(directory, recursive: true);
        }
    }

    [Fact]
    public void HtmlRenderCapabilityGallery_BindsInputPdfPreviewsDiagnosticsAndExpectations() {
        string html = "<style>" + CreatePortableEmbeddedFontFaceCss("Capability Gallery Test", 0x7E26, 0x66F8, 0x304D) + "</style>" + """
            <style>
            @page { size: 5in 4in; margin: 24px; @top-center { content: "Quarterly report" } }
            body { margin: 0; font: 11px/15px 'Capability Gallery Test'; }
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
                RenderOptions = new HtmlToPdfOptions {
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
