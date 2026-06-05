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

    /// <summary>Markdown reader options used by string and file overloads.</summary>
    public MarkdownReaderOptions? ReaderOptions { get; set; }

    /// <summary>Applies the built-in Word-like PDF theme before rendering Markdown blocks.</summary>
    public bool ApplyWordLikeTheme { get; set; } = true;

    private MarkdownPdfVisualTheme? _visualTheme;

    /// <summary>
    /// Markdown-aware visual theme. When null, the adapter uses the front matter <c>pdfTheme</c> value when present,
    /// otherwise falls back to the Word-like profile when <see cref="ApplyWordLikeTheme"/> is true.
    /// </summary>
    public MarkdownPdfVisualTheme? VisualTheme {
        get => _visualTheme?.Clone();
        set => _visualTheme = value?.Clone();
    }

    /// <summary>Use a <c>pdfTheme</c> or <c>pdf-theme</c> front matter value as the visual profile when no explicit visual theme is set.</summary>
    public bool UseFrontMatterVisualTheme { get; set; } = true;

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

    /// <summary>When true, supported local image files are embedded as PDF images.</summary>
    public bool IncludeLocalImages { get; set; } = true;

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

    /// <summary>Warnings recorded during the latest export.</summary>
    public IList<MarkdownPdfExportWarning> Warnings { get; } = new List<MarkdownPdfExportWarning>();

    internal void ResetExportState() {
        Warnings.Clear();
    }
}
