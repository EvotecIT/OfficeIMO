using System.Collections.Generic;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Options for exporting parser-supported PDFs to HTML through the first-party OfficeIMO logical PDF model.
/// </summary>
public sealed class PdfHtmlSaveOptions {
    /// <summary>
    /// HTML export profile. Defaults to semantic HTML.
    /// </summary>
    public PdfHtmlProfile Profile { get; set; } = PdfHtmlProfile.Semantic;

    /// <summary>
    /// Layout extraction options used when loading PDF bytes, paths, or streams into <see cref="PdfCore.PdfLogicalDocument"/>.
    /// </summary>
    public PdfCore.PdfTextLayoutOptions? LayoutOptions { get; set; }

    /// <summary>
    /// Optional selected source page ranges. When omitted, all pages are exported.
    /// </summary>
    public IReadOnlyList<PdfCore.PdfPageRange>? PageRanges { get; set; }

    /// <summary>
    /// Emit document metadata into the HTML head and body where useful.
    /// </summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>
    /// Emit page containers and page number metadata.
    /// </summary>
    public bool IncludePageContainers { get; set; } = true;

    /// <summary>
    /// Emit readable image output for image XObjects discovered in the logical model.
    /// </summary>
    public bool IncludeImagePlaceholders { get; set; } = true;

    /// <summary>
    /// Controls whether images are emitted as placeholders or embedded data URI image elements when extracted bytes are available.
    /// </summary>
    public PdfHtmlImageExportMode ImageExportMode { get; set; } = PdfHtmlImageExportMode.EmbeddedDataUri;

    /// <summary>
    /// Maximum extracted image byte length that may be embedded into generated HTML. Set to null to disable this guard.
    /// </summary>
    public long? MaxEmbeddedImageBytes { get; set; } = 10L * 1024L * 1024L;

    /// <summary>
    /// Emit link annotation placeholders. Semantic output emits a links section; positioned output emits positioned link frames.
    /// </summary>
    public bool IncludeLinkAnnotations { get; set; }

    /// <summary>
    /// Emit AcroForm widget placeholders. Semantic output emits a form-fields section; positioned output emits positioned form field frames.
    /// </summary>
    public bool IncludeFormWidgets { get; set; }

    /// <summary>
    /// Emit a complete HTML document with doctype, html, head, and body wrappers.
    /// </summary>
    public bool EmitDocumentShell { get; set; } = true;

    /// <summary>
    /// HTML document title used when PDF metadata does not provide one.
    /// </summary>
    public string DocumentTitleFallback { get; set; } = "OfficeIMO PDF Export";

    /// <summary>
    /// Shared conversion report populated by the HTML/PDF bridge.
    /// </summary>
    public PdfCore.PdfConversionReport ConversionReport { get; } = new PdfCore.PdfConversionReport();

    internal void ResetExportState() {
        ConversionReport.Clear();
    }
}
