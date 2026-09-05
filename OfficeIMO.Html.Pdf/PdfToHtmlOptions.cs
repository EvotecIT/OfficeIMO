using System.Collections.Generic;
using System.Linq;
using OfficeIMO.Drawing;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Html.Pdf;

/// <summary>
/// Options for exporting parser-supported PDFs to HTML through the first-party OfficeIMO logical PDF model.
/// </summary>
public sealed class PdfToHtmlOptions {
    /// <summary>Cancellation observed at page and export-summary boundaries.</summary>
    internal System.Threading.CancellationToken CancellationToken { get; set; }

    private OfficeHtmlDocumentOptions _documentOutput = new() {
        Title = "OfficeIMO PDF Export",
        Language = null,
        Theme = OfficeVisualThemeKind.Report,
        BodyClass = "officeimo-html officeimo-pdf-html"
    };

    /// <summary>Creates polished semantic PDF review HTML using the shared OfficeIMO document shell.</summary>
    public static PdfToHtmlOptions CreateSemanticProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.Report) => new() {
        Profile = PdfHtmlProfile.Semantic,
        Theme = theme,
        IncludeDefaultStyles = true
    };

    /// <summary>
    /// Creates positioned PDF review HTML with page geometry, images, links, and form widgets enabled.
    /// The output remains inert review HTML and never executes PDF actions or JavaScript.
    /// </summary>
    public static PdfToHtmlOptions CreatePositionedReviewProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.Report) => new() {
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

    /// <summary>
    /// Canonical semantic-read settings used when exporting an opened <see cref="PdfCore.PdfDocument"/>.
    /// Null uses <see cref="PdfCore.PdfReadOptions.Default"/>. <see cref="PageRanges"/> overrides the configured
    /// <see cref="PdfCore.PdfReadOptions.PageSelection"/> when both are supplied. This setting is ignored when the source
    /// is already a <see cref="PdfCore.PdfDocumentReadResult"/>.
    /// </summary>
    public PdfCore.PdfReadOptions? ReadOptions { get; set; }

    /// <summary>Composed document-versus-fragment, theme, title, language, style, and newline settings.</summary>
    public OfficeHtmlDocumentOptions DocumentOutput {
        get => _documentOutput;
        set => _documentOutput = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Compatibility alias for <see cref="OfficeHtmlDocumentOptions.Theme"/>.</summary>
    public OfficeVisualThemeKind Theme { get => DocumentOutput.Theme; set => DocumentOutput.Theme = value; }

    /// <summary>
    /// Emit the shared responsive OfficeIMO theme and adapter presentation styles.
    /// Positioned output always retains the minimum structural CSS required to preserve source geometry.
    /// </summary>
    public bool IncludeDefaultStyles { get => DocumentOutput.IncludeDefaultStyles; set => DocumentOutput.IncludeDefaultStyles = value; }

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
    /// Maximum UTF-16 characters retained by generated HTML. Set to null for the existing unbounded output behavior.
    /// </summary>
    public int? MaximumOutputCharacters { get; set; }

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
    public bool EmitDocumentShell { get => DocumentOutput.EmitDocumentShell; set => DocumentOutput.EmitDocumentShell = value; }

    /// <summary>
    /// HTML document title used when PDF metadata does not provide one.
    /// </summary>
    public string DocumentTitleFallback {
        get => DocumentOutput.Title ?? "OfficeIMO PDF Export";
        set => DocumentOutput.Title = value ?? throw new ArgumentNullException(nameof(value));
    }

    /// <summary>Language override for the HTML root. Null preserves the PDF catalog language.</summary>
    public string? Language { get => DocumentOutput.Language; set => DocumentOutput.Language = value; }

    /// <summary>Newline sequence used by generated HTML.</summary>
    public string NewLine { get => DocumentOutput.NewLine; set => DocumentOutput.NewLine = value; }

    internal PdfCore.PdfConversionReport Report { get; } = new PdfCore.PdfConversionReport();

    internal PdfToHtmlOptions CloneForConversion() => new() {
        CancellationToken = CancellationToken,
        Profile = Profile,
        ReadOptions = ReadOptions,
        DocumentOutput = DocumentOutput.Clone(),
        PageRanges = PageRanges?.ToArray(),
        UseSharedPageReadingOrder = UseSharedPageReadingOrder,
        IncludeMetadata = IncludeMetadata,
        IncludeOutlines = IncludeOutlines,
        IncludePageContainers = IncludePageContainers,
        IncludeImagePlaceholders = IncludeImagePlaceholders,
        ImageExportMode = ImageExportMode,
        MaxEmbeddedImageBytes = MaxEmbeddedImageBytes,
        MaximumOutputCharacters = MaximumOutputCharacters,
        IncludeLinkAnnotations = IncludeLinkAnnotations,
        IncludeFormWidgets = IncludeFormWidgets,
    };

    internal void Validate() {
        if (!Enum.IsDefined(typeof(PdfHtmlProfile), Profile)) {
            throw new ArgumentOutOfRangeException(nameof(Profile), Profile, "PDF HTML profile is not supported.");
        }
        DocumentOutput.Validate();
        if (MaxEmbeddedImageBytes.HasValue && MaxEmbeddedImageBytes.Value < 0L) {
            throw new ArgumentOutOfRangeException(nameof(MaxEmbeddedImageBytes), "Maximum embedded image bytes cannot be negative.");
        }
        if (MaximumOutputCharacters.HasValue && MaximumOutputCharacters.Value <= 0) {
            throw new ArgumentOutOfRangeException(nameof(MaximumOutputCharacters), "Maximum output characters must be positive.");
        }
    }
}
