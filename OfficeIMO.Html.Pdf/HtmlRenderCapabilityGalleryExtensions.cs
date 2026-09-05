using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html.Pdf;

/// <summary>Produces reviewable, hash-bound artifacts from the canonical managed HTML renderer.</summary>
public static class HtmlRenderCapabilityGalleryExtensions {
    /// <summary>
    /// Writes the exact HTML input, paged PDF, selected-page PNG/SVG previews, and deterministic JSON/Markdown manifests.
    /// </summary>
    public static HtmlCapabilityGalleryManifest SaveRenderCapabilityGallery(
        this HtmlConversionDocument document,
        string directoryPath,
        HtmlRenderCapabilityGalleryOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Gallery output path cannot be empty.", nameof(directoryPath));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (options.RenderOptions == null) throw new ArgumentException("Gallery render options cannot be null.", nameof(options));
        if (options.PreviewPageIndex < 0) throw new ArgumentOutOfRangeException(nameof(options.PreviewPageIndex));

        string directory = Path.GetFullPath(directoryPath);
        Directory.CreateDirectory(directory);
        string prefix = NormalizeFileName(options.Scenario.Id);
        var artifacts = new List<HtmlCapabilityGalleryArtifact>();

        string inputPath = Path.Combine(directory, prefix + ".input.html");
        artifacts.Add(HtmlCapabilityGalleryArtifact.WriteTextFile("source", "input-html", inputPath, "text/html", document.SourceHtml));

        byte[] pdf = document.ToPdf(new HtmlPdfSaveOptions(options.RenderOptions));
        string pdfPath = Path.Combine(directory, prefix + ".pdf");
        WriteBytes(pdfPath, pdf);
        artifacts.Add(HtmlCapabilityGalleryArtifact.FromFile("pdf", "paged-pdf", pdfPath, "application/pdf"));

        OfficeImageExportResult png = document.ExportImage(OfficeImageExportFormat.Png, options.RenderOptions, options.PreviewPageIndex);
        string pngPath = Path.Combine(directory, prefix + ".preview.png");
        WriteBytes(pngPath, png.Bytes);
        artifacts.Add(HtmlCapabilityGalleryArtifact.FromFile("preview-png", "page-preview", pngPath, "image/png"));

        OfficeImageExportResult svg = document.ExportImage(OfficeImageExportFormat.Svg, options.RenderOptions, options.PreviewPageIndex);
        string svgPath = Path.Combine(directory, prefix + ".preview.svg");
        WriteBytes(svgPath, svg.Bytes);
        artifacts.Add(HtmlCapabilityGalleryArtifact.FromFile("preview-svg", "searchable-vector-preview", svgPath, "image/svg+xml"));

        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, options.RenderOptions);
        var result = new HtmlCapabilityGalleryResult(options.Scenario, artifacts, rendered.Diagnostics);
        var manifest = new HtmlCapabilityGalleryManifest(
            result,
            document.ProfileContract.Profile,
            roundTripScore: null,
            document.ResourceManifest,
            options.Expectations);
        HtmlCapabilityGalleryArtifact.WriteTextFile(
            "manifest-json",
            "manifest-json",
            Path.Combine(directory, prefix + ".manifest.json"),
            "application/json",
            HtmlCapabilityGalleryManifestJsonWriter.ToJson(manifest));
        HtmlCapabilityGalleryArtifact.WriteTextFile(
            "manifest-markdown",
            "manifest-markdown",
            Path.Combine(directory, prefix + ".manifest.md"),
            "text/markdown",
            HtmlCapabilityGalleryManifestWriter.ToMarkdown(manifest));
        return manifest;
    }

    private static void WriteBytes(string path, byte[] bytes) {
        string? directory = Path.GetDirectoryName(path);
        if (!string.IsNullOrEmpty(directory)) Directory.CreateDirectory(directory);
        File.WriteAllBytes(path, bytes);
    }

    private static string NormalizeFileName(string value) {
        var builder = new StringBuilder(value.Length);
        foreach (char character in value) {
            builder.Append(char.IsLetterOrDigit(character) || character is '-' or '_' ? character : '-');
        }
        string normalized = builder.ToString().Trim('-');
        return normalized.Length == 0 ? "html-render-gallery" : normalized;
    }
}
