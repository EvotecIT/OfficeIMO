using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Html;
using System.Net.Http;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Options controlling HTML to Word conversion.
    /// </summary>
    public class HtmlToWordOptions {
        private HtmlConversionLimits _limits = HtmlConversionLimits.CreateTrustedProfile();
        /// <summary>
        /// Creates the default OfficeIMO HTML import profile.
        /// </summary>
        /// <returns>A new <see cref="HtmlToWordOptions"/> instance with the default compatibility-oriented settings.</returns>
        public static HtmlToWordOptions CreateOfficeIMOProfile() {
            var options = new HtmlToWordOptions {
                ImageProcessing = ImageProcessingMode.Embed,
                MaxTableCells = null
            };
            options.AllowedImageUriSchemes.Add(Uri.UriSchemeFile);
            return options;
        }

        /// <summary>
        /// Creates a bounded offline profile for untrusted HTML ingestion.
        /// </summary>
        /// <remarks>
        /// The profile keeps document-provided stylesheet links disabled, embeds only data URI images,
        /// blocks external image and stylesheet URI schemes, enables accessibility diagnostics, and
        /// applies conservative HTML, CSS, image, and table limits. Callers can relax individual
        /// limits or allow-lists when their ingestion boundary is more trusted.
        /// </remarks>
        /// <returns>A new <see cref="HtmlToWordOptions"/> instance configured for untrusted HTML.</returns>
        public static HtmlToWordOptions CreateUntrustedHtmlProfile() {
            var options = new HtmlToWordOptions {
                ImageProcessing = ImageProcessingMode.EmbedDataUriOnly,
                ResourceTimeout = TimeSpan.FromSeconds(5),
                MaxImageBytes = 5L * 1024L * 1024L,
                MaxTotalImageBytes = 20L * 1024L * 1024L,
                MaxHtmlNodes = 10000,
                MaxHtmlDepth = 64,
                MaxCssBytes = 256L * 1024L,
                MaxTotalCssBytes = 512L * 1024L,
                MaxTableCells = 50000,
                EnableAccessibilityDiagnostics = true,
                UnsupportedCssHandling = HtmlUnsupportedCssHandling.Warn
            };

            options.AllowedImageUriSchemes.Clear();
            options.AllowedImageUriSchemes.Add("data");
            options.AllowedStylesheetUriSchemes.Clear();

            return options;
        }

        /// <summary>
        /// Creates a profile for trusted HTML documents whose own linked stylesheets may be loaded.
        /// </summary>
        /// <remarks>
        /// This profile preserves default conversion behavior and validation settings while enabling
        /// <see cref="AllowDocumentStylesheetLinks"/>. Callers should still configure host allow-lists
        /// or byte limits when trusted documents can reference broad network locations.
        /// </remarks>
        /// <returns>A new <see cref="HtmlToWordOptions"/> instance configured for trusted document links.</returns>
        public static HtmlToWordOptions CreateTrustedDocumentProfile() {
            var options = new HtmlToWordOptions {
                ImageProcessing = ImageProcessingMode.Embed,
                AllowDocumentStylesheetLinks = true
            };
            options.AllowedImageUriSchemes.Add(Uri.UriSchemeFile);
            return options;
        }

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
        /// Shared URL policy applied before imported HTML anchors are materialized as Word hyperlinks.
        /// </summary>
        public HtmlUrlPolicy HyperlinkUrlPolicy { get; set; } = HtmlUrlPolicy.CreateHyperlinkProfile();

        /// <summary>
        /// Controls how images are processed during conversion.
        /// </summary>
        public ImageProcessingMode ImageProcessing { get; set; } = ImageProcessingMode.EmbedDataUriOnly;

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
        /// Optional maximum number of remote image candidates probed while selecting a source for one HTML image element.
        /// Defaults to one probe to avoid request fan-out from large srcset or picture candidate lists.
        /// Set to <see langword="null"/> to restore unbounded probing for trusted HTML.
        /// </summary>
        public int? MaxRemoteImageCandidateProbes { get; set; } = 1;

        /// <summary>
        /// Optional maximum number of responsive image candidates considered from <c>picture</c> and <c>srcset</c> inputs per image element.
        /// The direct <c>img src</c> fallback is still considered after this limit. Defaults to 32 candidates.
        /// Set to <see langword="null"/> to restore unbounded candidate expansion for trusted HTML.
        /// </summary>
        public int? MaxImageSourceCandidates { get; set; } = 32;

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
        /// Image URI schemes allowed during import. Defaults allow HTTP, HTTPS, and data URI images.
        /// Remote image embedding still requires <see cref="ImageProcessingMode.Embed"/> through an explicit option
        /// or compatibility profile.
        /// Add <see cref="Uri.UriSchemeFile"/> or use <see cref="CreateTrustedDocumentProfile"/> for trusted local-file images.
        /// Remove entries to reject matching image sources before they are loaded or linked.
        /// </summary>
        public HashSet<string> AllowedImageUriSchemes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            Uri.UriSchemeHttp,
            Uri.UriSchemeHttps,
            "data"
        };

        /// <summary>
        /// Optional host allow-list for absolute non-file image URIs. When empty, all hosts are allowed.
        /// </summary>
        public HashSet<string> AllowedImageHosts { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// Stylesheet URI schemes allowed during import. Defaults allow HTTP, HTTPS, and file-based stylesheets.
        /// Remove entries to reject matching stylesheet sources before they are loaded.
        /// </summary>
        public HashSet<string> AllowedStylesheetUriSchemes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            Uri.UriSchemeHttp,
            Uri.UriSchemeHttps,
            Uri.UriSchemeFile
        };

        /// <summary>
        /// Optional host allow-list for absolute non-file stylesheet URIs. When empty, all hosts are allowed.
        /// </summary>
        public HashSet<string> AllowedStylesheetHosts { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        /// <summary>
        /// When true, validates declared content types for remote stylesheet resources.
        /// Stylesheets with rejected content types are skipped and a diagnostic is emitted.
        /// </summary>
        public bool ValidateStylesheetContentTypes { get; set; } = true;

        /// <summary>
        /// Declared stylesheet media types allowed when <see cref="ValidateStylesheetContentTypes"/> is enabled.
        /// </summary>
        public HashSet<string> AllowedStylesheetContentTypes { get; } = new HashSet<string>(StringComparer.OrdinalIgnoreCase) {
            "text/css"
        };

        /// <summary>
        /// Optional maximum number of parsed HTML nodes allowed for a conversion operation.
        /// When exceeded, conversion stops with <see cref="HtmlConversionLimitException"/> and an error diagnostic.
        /// </summary>
        public int? MaxHtmlNodes {
            get => Limits.MaxHtmlNodes;
            set => Limits.MaxHtmlNodes = value;
        }

        /// <summary>
        /// Optional maximum parsed HTML tree depth allowed for a conversion operation.
        /// When exceeded, conversion stops with <see cref="HtmlConversionLimitException"/> and an error diagnostic.
        /// </summary>
        public int? MaxHtmlDepth {
            get => Limits.MaxHtmlDepth;
            set => Limits.MaxHtmlDepth = value;
        }

        /// <summary>
        /// Optional maximum UTF-8 byte count allowed for each stylesheet before parsing.
        /// When exceeded, conversion stops with <see cref="HtmlConversionLimitException"/> and an error diagnostic.
        /// </summary>
        public long? MaxCssBytes {
            get => Limits.MaxCssBytes;
            set => Limits.MaxCssBytes = value;
        }

        /// <summary>
        /// Optional maximum UTF-8 byte count allowed across all stylesheets in a single import operation.
        /// When exceeded, conversion stops with <see cref="HtmlConversionLimitException"/> and an error diagnostic.
        /// </summary>
        public long? MaxTotalCssBytes {
            get => Limits.MaxTotalCssBytes;
            set => Limits.MaxTotalCssBytes = value;
        }

        /// <summary>Shared HTML parsing and CSS limits used by the core engine and Word adapter.</summary>
        public HtmlConversionLimits Limits {
            get => _limits;
            set => _limits = value ?? HtmlConversionLimits.CreateTrustedProfile();
        }

        /// <summary>
        /// Optional maximum number of Word table cells allowed for a single imported HTML table.
        /// Defaults to 50,000 cells. Spans are resolved before the limit is checked. When exceeded, conversion stops with
        /// <see cref="HtmlConversionLimitException"/> and an error diagnostic.
        /// </summary>
        public long? MaxTableCells { get; set; } = 50000;

        internal HtmlDiagnosticReport ConversionReport { get; } = new HtmlDiagnosticReport();

        /// <summary>
        /// Optional conversion-scoped resolver for CSS classes without a built-in Word style mapping.
        /// This avoids global event state when conversions run concurrently.
        /// </summary>
        public Action<StyleMissingEventArgs>? StyleMissingHandler { get; set; }

        /// <summary>
        /// When true, emits advisory accessibility diagnostics for imported HTML patterns that can reduce
        /// document usability, such as missing image alternate text, weak link text, skipped heading levels,
        /// and data tables without header cells.
        /// </summary>
        public bool EnableAccessibilityDiagnostics { get; set; }

        /// <summary>
        /// When true, raw HTML comment nodes are imported as native Word comments anchored at their DOM position.
        /// Empty comments are skipped. Existing exported OfficeIMO comment sections are always imported through
        /// their linked comment references.
        /// </summary>
        public bool ImportHtmlComments { get; set; }

        /// <summary>
        /// Author name used for native Word comments created from raw HTML comment nodes.
        /// </summary>
        public string HtmlCommentAuthor { get; set; } = "HTML";

        /// <summary>
        /// Author initials used for native Word comments created from raw HTML comment nodes.
        /// </summary>
        public string HtmlCommentInitials { get; set; } = "HTML";

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

        /// <summary>
        /// Creates a copy of the current options instance so callers can reuse option templates safely.
        /// </summary>
        /// <remarks>
        /// Configuration values, allow-lists, and configured stylesheets are copied. Runtime diagnostics
        /// start empty on the clone so evidence from one conversion is never carried into the next.
        /// </remarks>
        /// <returns>A new <see cref="HtmlToWordOptions"/> with the same configuration values.</returns>
        public HtmlToWordOptions Clone() {
            var clone = new HtmlToWordOptions {
                FontFamily = FontFamily,
                QuotePrefix = QuotePrefix,
                QuoteSuffix = QuoteSuffix,
                DefaultPageSize = DefaultPageSize,
                DefaultOrientation = DefaultOrientation,
                IncludeListStyles = IncludeListStyles,
                ContinueNumbering = ContinueNumbering,
                SupportsHeadingNumbering = SupportsHeadingNumbering,
                BasePath = BasePath,
                NoteReferenceType = NoteReferenceType,
                LinkNoteUrls = LinkNoteUrls,
                HyperlinkUrlPolicy = HyperlinkUrlPolicy?.Clone() ?? HtmlUrlPolicy.CreateHyperlinkProfile(),
                ImageProcessing = ImageProcessing,
                HttpClient = HttpClient,
                ResourceTimeout = ResourceTimeout,
                MaxImageBytes = MaxImageBytes,
                MaxTotalImageBytes = MaxTotalImageBytes,
                MaxRemoteImageCandidateProbes = MaxRemoteImageCandidateProbes,
                MaxImageSourceCandidates = MaxImageSourceCandidates,
                ValidateImageContentTypes = ValidateImageContentTypes,
                ValidateStylesheetContentTypes = ValidateStylesheetContentTypes,
                Limits = Limits.Clone(),
                MaxTableCells = MaxTableCells,
                StyleMissingHandler = StyleMissingHandler,
                EnableAccessibilityDiagnostics = EnableAccessibilityDiagnostics,
                ImportHtmlComments = ImportHtmlComments,
                HtmlCommentAuthor = HtmlCommentAuthor,
                HtmlCommentInitials = HtmlCommentInitials,
                UnsupportedCssHandling = UnsupportedCssHandling,
                AllowDocumentStylesheetLinks = AllowDocumentStylesheetLinks,
                RenderPreAsTable = RenderPreAsTable,
                TableCaptionPosition = TableCaptionPosition,
                SectionTagHandling = SectionTagHandling
            };

            CopyDictionary(ClassStyles, clone.ClassStyles);
            CopyList(StylesheetPaths, clone.StylesheetPaths);
            CopyList(StylesheetContents, clone.StylesheetContents);
            CopySet(AllowedImageContentTypes, clone.AllowedImageContentTypes);
            CopySet(AllowedImageUriSchemes, clone.AllowedImageUriSchemes);
            CopySet(AllowedImageHosts, clone.AllowedImageHosts);
            CopySet(AllowedStylesheetUriSchemes, clone.AllowedStylesheetUriSchemes);
            CopySet(AllowedStylesheetHosts, clone.AllowedStylesheetHosts);
            CopySet(AllowedStylesheetContentTypes, clone.AllowedStylesheetContentTypes);

            return clone;
        }

        private static void CopyDictionary<TKey, TValue>(IDictionary<TKey, TValue> source, IDictionary<TKey, TValue> destination) {
            destination.Clear();
            foreach (var pair in source) {
                destination[pair.Key] = pair.Value;
            }
        }

        private static void CopyList<T>(IEnumerable<T> source, ICollection<T> destination) {
            destination.Clear();
            foreach (var item in source) {
                destination.Add(item);
            }
        }

        private static void CopySet<T>(IEnumerable<T> source, ISet<T> destination) {
            destination.Clear();
            foreach (var item in source) {
                destination.Add(item);
            }
        }
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
        /// Downloads and embeds all images into the document. Use asynchronous conversion when remote images are present.
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
