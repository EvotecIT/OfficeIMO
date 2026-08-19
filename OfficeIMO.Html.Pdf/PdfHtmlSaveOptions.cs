using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Options for exporting parser-supported PDFs to HTML through the first-party OfficeIMO logical PDF model.
/// </summary>
public sealed class PdfHtmlSaveOptions {
    /// <summary>Creates polished semantic PDF review HTML using the shared OfficeIMO document shell.</summary>
    public static PdfHtmlSaveOptions CreateSemanticProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.Report) => new() {
        Profile = PdfHtmlProfile.Semantic,
        Theme = theme,
        IncludeDefaultStyles = true
    };

    /// <summary>
    /// Creates positioned PDF review HTML with page geometry, images, links, and form widgets enabled.
    /// The output remains inert review HTML and never executes PDF actions or JavaScript.
    /// </summary>
    public static PdfHtmlSaveOptions CreatePositionedReviewProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.Report) => new() {
        Profile = PdfHtmlProfile.PositionedReview,
        Theme = theme,
        IncludeDefaultStyles = true,
        IncludeLinkAnnotations = true,
        IncludeFormWidgets = true
    };

    /// <summary>
    /// HTML export profile. Defaults to semantic HTML.
    /// </summary>
    public PdfHtmlProfile Profile { get; set; } = PdfHtmlProfile.Semantic;

    /// <summary>Shared OfficeIMO visual theme used by complete HTML document output.</summary>
    public OfficeVisualThemeKind Theme { get; set; } = OfficeVisualThemeKind.Report;

    /// <summary>
    /// Emit the shared responsive OfficeIMO theme and adapter presentation styles.
    /// Positioned output always retains the minimum structural CSS required to preserve source geometry.
    /// </summary>
    public bool IncludeDefaultStyles { get; set; } = true;

    /// <summary>
    /// Optional selected source page ranges. When omitted, all pages are exported.
    /// </summary>
    public IReadOnlyList<PdfCore.PdfPageRange>? PageRanges { get; set; }

    /// <summary>
    /// Whether semantic HTML should use the shared crop-, rotation-, and column-aware logical reading order.
    /// Positioned review output always retains source geometry.
    /// </summary>
    public bool UseSharedPageReadingOrder { get; set; } = true;

    /// <summary>
    /// Emit document metadata into the HTML head and body where useful.
    /// </summary>
    public bool IncludeMetadata { get; set; } = true;

    /// <summary>
    /// Emit PDF outlines/bookmarks as inert HTML navigation metadata when available.
    /// </summary>
    public bool IncludeOutlines { get; set; } = true;

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

    internal PdfCore.PdfConversionReport Report { get; } = new PdfCore.PdfConversionReport();

    internal PdfHtmlSaveOptions CloneForConversion() => new() {
        Profile = Profile,
        Theme = Theme,
        IncludeDefaultStyles = IncludeDefaultStyles,
        PageRanges = PageRanges?.ToArray(),
        UseSharedPageReadingOrder = UseSharedPageReadingOrder,
        IncludeMetadata = IncludeMetadata,
        IncludeOutlines = IncludeOutlines,
        IncludePageContainers = IncludePageContainers,
        IncludeImagePlaceholders = IncludeImagePlaceholders,
        ImageExportMode = ImageExportMode,
        MaxEmbeddedImageBytes = MaxEmbeddedImageBytes,
        IncludeLinkAnnotations = IncludeLinkAnnotations,
        IncludeFormWidgets = IncludeFormWidgets,
        EmitDocumentShell = EmitDocumentShell,
        DocumentTitleFallback = DocumentTitleFallback
    };

    internal void Validate() {
        if (!Enum.IsDefined(typeof(PdfHtmlProfile), Profile)) {
            throw new ArgumentOutOfRangeException(nameof(Profile), Profile, "PDF HTML profile is not supported.");
        }
        if (IncludeDefaultStyles && !Enum.IsDefined(typeof(OfficeVisualThemeKind), Theme)) {
            throw new ArgumentOutOfRangeException(nameof(Theme), Theme, "Office HTML theme is not supported.");
        }
        if (MaxEmbeddedImageBytes.HasValue && MaxEmbeddedImageBytes.Value < 0L) {
            throw new ArgumentOutOfRangeException(nameof(MaxEmbeddedImageBytes), "Maximum embedded image bytes cannot be negative.");
        }
    }
}
