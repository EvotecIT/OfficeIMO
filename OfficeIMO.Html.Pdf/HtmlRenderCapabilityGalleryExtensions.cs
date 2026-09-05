using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Linq;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.Core.Internal;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>Produces reviewable, hash-bound artifacts from the canonical managed HTML renderer.</summary>
public static class HtmlRenderCapabilityGalleryExtensions {
    /// <summary>
    /// Writes the exact HTML input, paged PDF, selected or all-page previews, and manifests with artifact-specific evidence.
    /// </summary>
    public static HtmlCapabilityGalleryManifest SaveRenderCapabilityGallery(
        this HtmlConversionDocument document,
        string directoryPath,
        HtmlRenderCapabilityGalleryOptions options) {
        if (document == null) throw new ArgumentNullException(nameof(document));
        if (string.IsNullOrWhiteSpace(directoryPath)) throw new ArgumentException("Gallery output path cannot be empty.", nameof(directoryPath));
        if (options == null) throw new ArgumentNullException(nameof(options));
        if (options.RenderOptions == null) throw new ArgumentException("Gallery render options cannot be null.", nameof(options));
        if (options.PdfProofOptions == null) throw new ArgumentException("Gallery PDF proof options cannot be null.", nameof(options));
        if (options.PreviewPageIndex < 0) throw new ArgumentOutOfRangeException(nameof(options.PreviewPageIndex));
        OfficeImageExportFormat[] formats = options.PreviewFormats.Distinct().ToArray();
        if (formats.Length == 0 || formats.Any(format => !format.IsRaster() && format != OfficeImageExportFormat.Svg))
            throw new ArgumentException("At least one supported preview format is required.", nameof(options));

        HtmlToPdfOptions pdfOptions = options.RenderOptions.ClonePdf();
        HtmlRenderOptions renderOptions = HtmlPdfRenderedConverter.ResolveRenderOptions(pdfOptions);
        HtmlRenderDocument rendered = HtmlRenderEngine.Render(document, renderOptions);
        if (!options.PreviewAllPages && options.PreviewPageIndex >= rendered.Pages.Count)
            throw new ArgumentOutOfRangeException(nameof(options.PreviewPageIndex), "The selected preview page does not exist.");

        string directory = Path.GetFullPath(directoryPath);
        Directory.CreateDirectory(directory);
        string prefix = NormalizeFileName(options.Scenario.Id);
        var artifacts = new List<HtmlCapabilityGalleryArtifact>();

        string inputPath = Path.Combine(directory, prefix + ".input.html");
        artifacts.Add(HtmlCapabilityGalleryArtifact.WriteTextFile("source", "input-html", inputPath, "text/html", document.SourceHtml));

        PdfCore.PdfDocumentConversionResult conversion = HtmlPdfConverterExtensions.CreateResult(
            HtmlPdfRenderedConverter.CreatePdf(rendered, pdfOptions, CancellationToken.None));
        byte[] pdf = conversion.ToBytes();
        string pdfPath = Path.Combine(directory, prefix + ".pdf");
        WriteBytes(pdfPath, pdf);
        PdfCore.PdfConversionProofReport proof = conversion.AssessArtifactProof(pdf, options.PdfProofOptions);
        HtmlDiagnostic[] pdfDiagnostics = conversion.Warnings.Select(warning => new HtmlDiagnostic(
            warning.Converter, warning.Code, warning.Message,
            warning.Severity == PdfCore.PdfConversionWarningSeverity.Error ? HtmlDiagnosticSeverity.Error :
                warning.Severity == PdfCore.PdfConversionWarningSeverity.Warning ? HtmlDiagnosticSeverity.Warning : HtmlDiagnosticSeverity.Info,
            warning.Source, lossKind: warning.LossKind)).ToArray();
        artifacts.Add(HtmlCapabilityGalleryArtifact.FromFile("pdf", "paged-pdf", pdfPath, "application/pdf")
            .WithEvidence(new HtmlCapabilityGalleryArtifactEvidence(
                proof.DocumentInfo?.PageCount ?? rendered.Pages.Count, null, null, null, "pt", pdfDiagnostics,
                new[] { new HtmlCapabilityGalleryCheck("pdf-proof", proof.IsSatisfied, proof.Summary) })));

        IEnumerable<int> pages = options.PreviewAllPages
            ? Enumerable.Range(0, rendered.Pages.Count) : new[] { options.PreviewPageIndex };
        foreach (int pageIndex in pages) {
            foreach (OfficeImageExportFormat format in formats) {
                OfficeImageExportResult image = HtmlImageExportExtensions.RenderPage(
                    rendered.Pages[pageIndex], format, renderOptions, rendered.DiagnosticReport, CancellationToken.None);
                string extension = format.GetFileExtension().TrimStart('.');
                string suffix = options.PreviewAllPages ? ".page-" + (pageIndex + 1).ToString("D4", System.Globalization.CultureInfo.InvariantCulture) : ".preview";
                string imagePath = Path.Combine(directory, prefix + suffix + "." + extension);
                WriteBytes(imagePath, image.Bytes);
                HtmlDiagnostic[] imageDiagnostics = image.Diagnostics.Select(diagnostic => new HtmlDiagnostic(
                    "OfficeIMO.Html.Image", diagnostic.Code, diagnostic.Message,
                    diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Error ? HtmlDiagnosticSeverity.Error :
                        diagnostic.Severity == OfficeImageExportDiagnosticSeverity.Warning ? HtmlDiagnosticSeverity.Warning : HtmlDiagnosticSeverity.Info,
                    diagnostic.Source, lossKind: diagnostic.LossKind)).ToArray();
                artifacts.Add(HtmlCapabilityGalleryArtifact.FromFile(
                    (options.PreviewAllPages ? "page-" + (pageIndex + 1) + "-" : "preview-") + extension,
                    format == OfficeImageExportFormat.Svg ? "searchable-vector-preview" : "page-preview", imagePath, format.GetMimeType())
                    .WithEvidence(new HtmlCapabilityGalleryArtifactEvidence(
                        rendered.Pages.Count, pageIndex + 1, image.Width, image.Height, "px", imageDiagnostics,
                        new[] { new HtmlCapabilityGalleryCheck("image-content", true, "Encoded format and dimensions validated by OfficeImageExportResult.") })));
            }
        }
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
        OfficeFileCommit.WriteAllBytes(path, bytes);
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
