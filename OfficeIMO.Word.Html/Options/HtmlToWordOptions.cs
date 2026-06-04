using DocumentFormat.OpenXml.Wordprocessing;
using System.Net.Http;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Options controlling HTML to Word conversion.
    /// </summary>
    public class HtmlToWordOptions {
        /// <summary>
        /// Optional font family applied to created runs during conversion.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>
        /// Character inserted before inline quoted text. Defaults to left double quotation mark.
        /// </summary>
        public string QuotePrefix { get; set; } = "\u201C";

        /// <summary>
        /// Character inserted after inline quoted text. Defaults to right double quotation mark.
        /// </summary>
        public string QuoteSuffix { get; set; } = "\u201D";

        /// <summary>
        /// Optional default page size applied when creating new documents.
        /// </summary>
        public WordPageSize? DefaultPageSize { get; set; }

        /// <summary>
        /// Optional default page orientation applied when creating new documents.
        /// </summary>
        public PageOrientationValues? DefaultOrientation { get; set; }

        /// <summary>
        /// Maps HTML class names to paragraph styles. Example: <code>ClassStyles["title"] = WordParagraphStyles.Heading1;</code>
        /// </summary>
        public Dictionary<string, WordParagraphStyles> ClassStyles { get; } = new Dictionary<string, WordParagraphStyles>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// When true, attempts to include list styling information during conversion.
        /// </summary>
        public bool IncludeListStyles { get; set; }

        /// <summary>
        /// When true, numbered lists will continue numbering across separate lists.
        /// </summary>
        public bool ContinueNumbering { get; set; }

        /// <summary>
        /// When true, heading elements are converted into a numbered list using
        /// <see cref="WordListStyle.Headings111"/> so headings receive automatic numbering.
        /// </summary>
        public bool SupportsHeadingNumbering { get; set; }

        /// <summary>
        /// Base directory used to resolve relative resource paths like images.
        /// </summary>
        public string? BasePath { get; set; }

        /// <summary>
        /// Controls whether HTML-generated notes are inserted as footnotes or endnotes.
        /// </summary>
        public NoteReferenceType NoteReferenceType { get; set; } = NoteReferenceType.Footnote;

        /// <summary>
        /// When true, URLs used as note text are emitted as hyperlinks inside the note.
        /// </summary>
        public bool LinkNoteUrls { get; set; } = true;

        /// <summary>
        /// Controls how images are processed during conversion.
        /// </summary>
        public ImageProcessingMode ImageProcessing { get; set; } = ImageProcessingMode.Embed;

        /// <summary>
        /// Optional <see cref="HttpClient"/> used to download remote resources (images, SVG).
        /// If not provided, a shared client instance is used.
        /// </summary>
        public HttpClient? HttpClient { get; set; }

        /// <summary>
        /// Optional timeout applied when downloading remote resources.
        /// </summary>
        public TimeSpan? ResourceTimeout { get; set; }

        /// <summary>
        /// Optional maximum number of bytes allowed for a single image resource, including SVG images.
        /// When exceeded, the image is skipped, alt text is inserted when available, and a diagnostic is emitted.
        /// </summary>
        public long? MaxImageBytes { get; set; }

        /// <summary>
        /// Optional maximum number of image bytes allowed across a single HTML import operation, including SVG images.
        /// When exceeded, the image that crosses the budget is skipped, alt text is inserted when available, and a diagnostic is emitted.
        /// </summary>
        public long? MaxTotalImageBytes { get; set; }

        /// <summary>
        /// When true, validates declared image content types for remote image resources and data URI images.
        /// Images with rejected content types are skipped, alt text is inserted when available, and a diagnostic is emitted.
        /// </summary>
        public bool ValidateImageContentTypes { get; set; } = true;

        /// <summary>
        /// Declared image media types allowed when <see cref="ValidateImageContentTypes"/> is enabled.
        /// Add <c>image/*</c> to allow any declared image media type.
        /// </summary>
        public HashSet<string> AllowedImageContentTypes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "image/png",
            "image/jpeg",
            "image/jpg",
            "image/gif",
            "image/bmp",
            "image/tiff",
            "image/webp",
            "image/svg+xml"
        };

        /// <summary>
        /// Image URI schemes allowed during import. Defaults allow HTTP, HTTPS, file, and data URI images.
        /// Remove entries to reject matching image sources before they are loaded or linked.
        /// </summary>
        public HashSet<string> AllowedImageUriSchemes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            Uri.UriSchemeHttp,
            Uri.UriSchemeHttps,
            Uri.UriSchemeFile,
            "data"
        };

        /// <summary>
        /// Optional host allow-list for absolute non-file image URIs. When empty, all hosts are allowed.
        /// </summary>
        public HashSet<string> AllowedImageHosts { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Optional maximum number of parsed HTML nodes allowed for a conversion operation.
        /// When exceeded, conversion stops with <see cref="HtmlConversionLimitException"/> and an error diagnostic.
        /// </summary>
        public int? MaxHtmlNodes { get; set; }

        /// <summary>
        /// Optional maximum parsed HTML tree depth allowed for a conversion operation.
        /// When exceeded, conversion stops with <see cref="HtmlConversionLimitException"/> and an error diagnostic.
        /// </summary>
        public int? MaxHtmlDepth { get; set; }

        /// <summary>
        /// Optional maximum UTF-8 byte count allowed for each stylesheet before parsing.
        /// When exceeded, conversion stops with <see cref="HtmlConversionLimitException"/> and an error diagnostic.
        /// </summary>
        public long? MaxCssBytes { get; set; }

        /// <summary>
        /// Optional maximum number of Word table cells allowed for a single imported HTML table.
        /// Spans are resolved before the limit is checked. When exceeded, conversion stops with
        /// <see cref="HtmlConversionLimitException"/> and an error diagnostic.
        /// </summary>
        public long? MaxTableCells { get; set; }

        /// <summary>
        /// Diagnostics produced while converting HTML. The converter appends warnings here when
        /// content is skipped or degraded, such as an image that cannot be loaded.
        /// </summary>
        public List<HtmlConversionDiagnostic> Diagnostics { get; } = new List<HtmlConversionDiagnostic>();

        /// <summary>
        /// Optional callback invoked whenever a conversion diagnostic is produced.
        /// </summary>
        public Action<HtmlConversionDiagnostic>? DiagnosticHandler { get; set; }

        /// <summary>
        /// When true, emits advisory accessibility diagnostics for imported HTML patterns that can reduce
        /// document usability, such as missing image alternate text, weak link text, skipped heading levels,
        /// and data tables without header cells.
        /// </summary>
        public bool EnableAccessibilityDiagnostics { get; set; }

        /// <summary>
        /// Controls how unsupported CSS properties and values are handled during import.
        /// Defaults to warning diagnostics while preserving best-effort conversion.
        /// </summary>
        public HtmlUnsupportedCssHandling UnsupportedCssHandling { get; set; } = HtmlUnsupportedCssHandling.Warn;

        /// <summary>
        /// File paths pointing to external stylesheets that should be applied during conversion.
        /// </summary>
        public List<string> StylesheetPaths { get; } = new List<string>();

        /// <summary>
        /// Raw CSS stylesheet contents that should be applied during conversion.
        /// </summary>
        public List<string> StylesheetContents { get; } = new List<string>();

        /// <summary>
        /// When true, stylesheet links declared in the imported HTML document are loaded.
        /// Keep this disabled for untrusted HTML and prefer <see cref="StylesheetPaths"/> or
        /// <see cref="StylesheetContents"/> for caller-provided stylesheets.
        /// </summary>
        public bool AllowDocumentStylesheetLinks { get; set; }

        /// <summary>
        /// When true, <c>&lt;pre&gt;</c> elements are rendered inside a single-cell table.
        /// </summary>
        public bool RenderPreAsTable { get; set; }

        /// <summary>
        /// Specifies where table captions should be inserted relative to the table.
        /// </summary>
        public TableCaptionPosition TableCaptionPosition { get; set; } = TableCaptionPosition.Above;

        /// <summary>
        /// Controls how the <c>&lt;section&gt;</c> tag is mapped into Word.
        /// </summary>
        public SectionTagHandling SectionTagHandling { get; set; } = SectionTagHandling.WordSection;
    }

    /// <summary>
    /// Determines the position of a table caption relative to the table.
    /// </summary>
    public enum TableCaptionPosition {
        /// <summary>
        /// Caption is placed before the table.
        /// </summary>
        Above,

        /// <summary>
        /// Caption is placed after the table.
        /// </summary>
        Below
    }

    /// <summary>
    /// Specifies how images should be processed during HTML to Word conversion.
    /// </summary>
    public enum ImageProcessingMode {
        /// <summary>
        /// Downloads and embeds all images into the document (default).
        /// </summary>
        Embed,

        /// <summary>
        /// Links to external images via relationships instead of embedding them.
        /// Data URI images are still embedded.
        /// </summary>
        LinkExternal,

        /// <summary>
        /// Only embeds data URI images; external images are skipped.
        /// </summary>
        EmbedDataUriOnly
    }

    /// <summary>
    /// Specifies where generated note references are placed.
    /// </summary>
    public enum NoteReferenceType {
        /// <summary>
        /// Use footnotes (default).
        /// </summary>
        Footnote,

        /// <summary>
        /// Use endnotes.
        /// </summary>
        Endnote
    }

    /// <summary>
    /// Determines how the <c>&lt;section&gt;</c> HTML tag is represented in Word.
    /// </summary>
    public enum SectionTagHandling {
        /// <summary>
        /// Creates a new Word section for each <c>&lt;section&gt;</c>.
        /// </summary>
        WordSection,

        /// <summary>
        /// Treats <c>&lt;section&gt;</c> like a generic block container.
        /// </summary>
        Block
    }

    /// <summary>
    /// Severity for an HTML conversion diagnostic.
    /// </summary>
    public enum HtmlConversionDiagnosticSeverity {
        /// <summary>
        /// Informational diagnostic that does not indicate content loss.
        /// </summary>
        Info,

        /// <summary>
        /// Warning diagnostic for skipped or degraded content.
        /// </summary>
        Warning,

        /// <summary>
        /// Error diagnostic for content that could not be converted as requested.
        /// </summary>
        Error
    }

    /// <summary>
    /// Determines how the HTML importer handles CSS declarations it cannot map to Word output.
    /// </summary>
    public enum HtmlUnsupportedCssHandling {
        /// <summary>
        /// Unsupported CSS declarations are ignored without diagnostics.
        /// </summary>
        Ignore,

        /// <summary>
        /// Unsupported CSS declarations produce warning diagnostics while conversion continues.
        /// </summary>
        Warn,

        /// <summary>
        /// Unsupported CSS declarations produce error diagnostics and stop conversion.
        /// </summary>
        Error
    }

    /// <summary>
    /// Describes skipped, degraded, or otherwise notable content observed during HTML conversion.
    /// </summary>
    public sealed class HtmlConversionDiagnostic {
        /// <summary>
        /// Creates a conversion diagnostic.
        /// </summary>
        /// <param name="code">Stable diagnostic code.</param>
        /// <param name="message">Human-readable message.</param>
        /// <param name="severity">Diagnostic severity.</param>
        /// <param name="source">Optional HTML/resource source associated with the diagnostic.</param>
        /// <param name="detail">Optional low-level detail, such as an exception type or status text.</param>
        public HtmlConversionDiagnostic(string code, string message, HtmlConversionDiagnosticSeverity severity = HtmlConversionDiagnosticSeverity.Warning, string? source = null, string? detail = null) {
            Code = code ?? throw new ArgumentNullException(nameof(code));
            Message = message ?? throw new ArgumentNullException(nameof(message));
            Severity = severity;
            Source = source;
            Detail = detail;
        }

        /// <summary>
        /// Stable diagnostic code.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Human-readable diagnostic message.
        /// </summary>
        public string Message { get; }

        /// <summary>
        /// Diagnostic severity.
        /// </summary>
        public HtmlConversionDiagnosticSeverity Severity { get; }

        /// <summary>
        /// Optional HTML/resource source associated with the diagnostic.
        /// </summary>
        public string? Source { get; }

        /// <summary>
        /// Optional low-level detail, such as an exception type or status text.
        /// </summary>
        public string? Detail { get; }
    }

    /// <summary>
    /// Thrown when HTML conversion input exceeds a configured safety or resource limit.
    /// </summary>
    public sealed class HtmlConversionLimitException : InvalidOperationException {
        /// <summary>
        /// Creates a conversion limit exception.
        /// </summary>
        /// <param name="code">Stable diagnostic code associated with the limit.</param>
        /// <param name="message">Human-readable message.</param>
        /// <param name="source">Configured limit or resource source that was exceeded.</param>
        /// <param name="actual">Observed value.</param>
        /// <param name="limit">Configured limit.</param>
        /// <param name="detail">Optional formatted detail.</param>
        public HtmlConversionLimitException(string code, string message, string source, long actual, long limit, string? detail = null) : base(message) {
            Code = code ?? throw new ArgumentNullException(nameof(code));
            LimitSource = source ?? throw new ArgumentNullException(nameof(source));
            Actual = actual;
            Limit = limit;
            Detail = detail;
        }

        /// <summary>
        /// Stable diagnostic code associated with the limit.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// Configured limit or resource source that was exceeded.
        /// </summary>
        public string LimitSource { get; }

        /// <summary>
        /// Observed value.
        /// </summary>
        public long Actual { get; }

        /// <summary>
        /// Configured limit.
        /// </summary>
        public long Limit { get; }

        /// <summary>
        /// Optional formatted detail.
        /// </summary>
        public string? Detail { get; }
    }

    /// <summary>
    /// Thrown when unsupported CSS is encountered while strict CSS handling is enabled.
    /// </summary>
    public sealed class HtmlUnsupportedCssException : InvalidOperationException {
        /// <summary>
        /// Creates an unsupported CSS exception.
        /// </summary>
        /// <param name="code">Stable diagnostic code associated with the unsupported CSS declaration.</param>
        /// <param name="message">Human-readable message.</param>
        /// <param name="source">HTML element and CSS property associated with the unsupported declaration.</param>
        /// <param name="detail">Optional unsupported value detail.</param>
        public HtmlUnsupportedCssException(string code, string message, string source, string? detail = null) : base(message) {
            Code = code ?? throw new ArgumentNullException(nameof(code));
            CssSource = source ?? throw new ArgumentNullException(nameof(source));
            Detail = detail;
        }

        /// <summary>
        /// Stable diagnostic code associated with the unsupported CSS declaration.
        /// </summary>
        public string Code { get; }

        /// <summary>
        /// HTML element and CSS property associated with the unsupported declaration.
        /// </summary>
        public string CssSource { get; }

        /// <summary>
        /// Optional unsupported value detail.
        /// </summary>
        public string? Detail { get; }
    }
}
