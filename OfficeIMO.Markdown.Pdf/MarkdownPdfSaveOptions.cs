using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

/// <summary>
/// Options for converting OfficeIMO.Markdown documents to first-party PDF documents.
/// </summary>
public sealed class MarkdownPdfSaveOptions {
    private double _defaultImageWidth = 320D;
    private double _defaultImageHeight = 180D;
    private int _maximumDataUriImageBytes = 5 * 1024 * 1024;
    private int _maximumRemoteImageBytes = 5 * 1024 * 1024;

    /// <summary>PDF creation options passed to the first-party PDF engine.</summary>
    public PdfCore.PdfOptions? PdfOptions { get; set; }

    /// <summary>Optional Markdown default font family used by the first-party PDF engine.</summary>
    public string? FontFamily { get; set; }

    /// <summary>
    /// Built-in generated-text fallback groups applied by the Markdown PDF converter. Set to <see cref="PdfCore.PdfTextFallbackFeatures.None"/> for exact standard-font output.
    /// </summary>
    public PdfCore.PdfTextFallbackFeatures TextFallbacks { get; set; } = PdfCore.PdfTextFallbackFeatures.Default;

    /// <summary>
    /// When true, Markdown PDF export may load installed system fonts to cover Unicode, symbol, and emoji fallback runs.
    /// Defaults to true so Markdown text can render common Unicode content without caller preprocessing.
    /// </summary>
    public bool AllowSystemFontEmbedding { get; set; } = true;

    /// <summary>Applies the built-in Word-like PDF theme before rendering Markdown blocks.</summary>
    public bool ApplyDefaultTheme { get; set; } = true;

    private MarkdownVisualTheme? _theme;
    private MarkdownPdfStyle? _style;

    /// <summary>
    /// Shared Markdown visual theme used to align PDF output with HTML and Word exporters.
    /// When set, this is translated into PDF-specific styling unless <see cref="Style"/> is also set.
    /// </summary>
    public MarkdownVisualTheme? Theme {
        get => _theme?.Clone();
        set => _theme = value?.Clone();
    }

    internal MarkdownVisualTheme? ThemeSnapshot => _theme?.Clone();

    /// <summary>
    /// PDF-specific layout and rendering style overrides. Preset selection belongs to <see cref="Theme"/> so
    /// HTML, PDF, and Word share one visual-theme catalog.
    /// </summary>
    public MarkdownPdfStyle? Style {
        get => _style?.Clone();
        set => _style = value?.Clone();
    }

    /// <summary>Use <c>theme</c>, <c>visualTheme</c>, <c>pdfTheme</c>, or equivalent front matter values as the visual profile when no explicit theme is set.</summary>
    public bool UseFrontMatterTheme { get; set; } = true;

    /// <summary>Create PDF outlines from supported Markdown headings.</summary>
    public bool CreateOutlineFromHeadings { get; set; } = true;

    /// <summary>Use common front matter keys such as title, author, subject, description, keywords, and tags as PDF metadata.</summary>
    public bool UseFrontMatterMetadata { get; set; } = true;

    /// <summary>Controls whether front matter is rendered as a document heading, a metadata table, or hidden from the PDF body.</summary>
    public MarkdownPdfFrontMatterRenderMode FrontMatterRenderMode { get; set; } = MarkdownPdfFrontMatterRenderMode.DocumentHeader;

    /// <summary>Use the first Markdown heading as PDF title metadata when no explicit or front matter title is available.</summary>
    public bool UseFirstHeadingAsTitle { get; set; } = true;

    /// <summary>Explicit PDF title metadata. When null, front matter or the first heading can be used.</summary>
    public string? Title { get; set; }

    /// <summary>Explicit PDF author metadata. When null, front matter author can be used.</summary>
    public string? Author { get; set; }

    /// <summary>Explicit PDF subject metadata. When null, front matter subject, description, or summary can be used.</summary>
    public string? Subject { get; set; }

    /// <summary>Explicit PDF keywords metadata. When null, front matter keywords or tags can be used.</summary>
    public string? Keywords { get; set; }

    /// <summary>Base directory used to resolve relative local image paths.</summary>
    public string? BaseDirectory { get; set; }

    /// <summary>When true, supported local image files are embedded as PDF images. Defaults to false for untrusted Markdown.</summary>
    public bool IncludeLocalImages { get; set; }

    /// <summary>When true and <see cref="BaseDirectory"/> is set, local image paths must resolve inside that directory.</summary>
    public bool RestrictLocalImagesToBaseDirectory { get; set; } = true;

    /// <summary>When true, supported base64 data URI images are embedded as PDF images.</summary>
    public bool IncludeDataUriImages { get; set; } = true;

    /// <summary>Maximum decoded byte length for a single data URI image.</summary>
    public int MaximumDataUriImageBytes {
        get => _maximumDataUriImageBytes;
        set {
            if (value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(value), "Maximum data URI image bytes must be positive.");
            }

            _maximumDataUriImageBytes = value;
        }
    }

    /// <summary>
    /// Optional resolver for HTTPS/HTTP Markdown images. When null, remote images are rendered as placeholders and recorded as warnings.
    /// </summary>
    public Func<Uri, byte[]?>? RemoteImageResolver { get; set; }

    /// <summary>Maximum byte length accepted from <see cref="RemoteImageResolver"/> for a single remote image.</summary>
    public int MaximumRemoteImageBytes {
        get => _maximumRemoteImageBytes;
        set {
            if (value <= 0) {
                throw new ArgumentOutOfRangeException(nameof(value), "Maximum remote image bytes must be positive.");
            }

            _maximumRemoteImageBytes = value;
        }
    }

    /// <summary>Fallback image width in points when Markdown did not provide a width hint.</summary>
    public double DefaultImageWidth {
        get => _defaultImageWidth;
        set {
            if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new ArgumentOutOfRangeException(nameof(value), "Default image width must be a positive finite value.");
            }

            _defaultImageWidth = value;
        }
    }

    /// <summary>Fallback image height in points when Markdown did not provide a height hint.</summary>
    public double DefaultImageHeight {
        get => _defaultImageHeight;
        set {
            if (value <= 0 || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new ArgumentOutOfRangeException(nameof(value), "Default image height must be a positive finite value.");
            }

            _defaultImageHeight = value;
        }
    }

    /// <summary>
    /// Applies a high-level export profile by setting the Markdown PDF options that correspond to that profile.
    /// </summary>
    public MarkdownPdfSaveOptions UseProfile(PdfCore.PdfExportProfile profile) {
        switch (profile) {
            case PdfCore.PdfExportProfile.Faithful:
                IncludeDataUriImages = true;
                IncludeLocalImages = true;
                ApplyDefaultTheme = true;
                CreateOutlineFromHeadings = true;
                FrontMatterRenderMode = MarkdownPdfFrontMatterRenderMode.DocumentHeader;
                break;
            case PdfCore.PdfExportProfile.Lightweight:
                IncludeDataUriImages = false;
                IncludeLocalImages = false;
                RemoteImageResolver = null;
                ApplyDefaultTheme = true;
                CreateOutlineFromHeadings = true;
                break;
            case PdfCore.PdfExportProfile.PrintReady:
                IncludeDataUriImages = true;
                ApplyDefaultTheme = true;
                CreateOutlineFromHeadings = true;
                FrontMatterRenderMode = MarkdownPdfFrontMatterRenderMode.DocumentHeader;
                break;
            case PdfCore.PdfExportProfile.TextOnly:
                IncludeDataUriImages = false;
                IncludeLocalImages = false;
                RemoteImageResolver = null;
                ApplyDefaultTheme = false;
                FrontMatterRenderMode = MarkdownPdfFrontMatterRenderMode.Hidden;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(profile), profile, "Unsupported PDF export profile.");
        }

        return this;
    }

    /// <summary>Warnings recorded during the latest export.</summary>
    internal IList<MarkdownPdfExportWarning> Warnings { get; private set; } = new List<MarkdownPdfExportWarning>();

    /// <summary>
    /// Shared conversion report populated alongside <see cref="Warnings"/> for wrapper-friendly diagnostics.
    /// The report is cleared at the start of each export.
    /// </summary>
    internal PdfCore.PdfConversionReport Report { get; private set; } = new PdfCore.PdfConversionReport();

    internal MarkdownPdfSaveOptions CloneForConversion() {
        var clone = (MarkdownPdfSaveOptions)MemberwiseClone();
        clone._theme = _theme?.Clone();
        clone._style = _style?.Clone();
        clone.Warnings = new List<MarkdownPdfExportWarning>();
        clone.Report = new PdfCore.PdfConversionReport();
        return clone;
    }

    /// <summary>Creates an independent copy of these conversion options.</summary>
    public MarkdownPdfSaveOptions Clone() => CloneForConversion();
}
