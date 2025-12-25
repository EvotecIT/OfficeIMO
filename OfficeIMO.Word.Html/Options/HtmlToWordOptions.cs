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
        /// File paths pointing to external stylesheets that should be applied during conversion.
        /// </summary>
        public List<string> StylesheetPaths { get; } = new List<string>();

        /// <summary>
        /// Raw CSS stylesheet contents that should be applied during conversion.
        /// </summary>
        public List<string> StylesheetContents { get; } = new List<string>();

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
}
