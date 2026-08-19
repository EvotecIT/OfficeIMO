using OfficeIMO.Drawing;
using OfficeIMO.Html;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Options controlling Word to HTML conversion.
    /// </summary>
    public class WordToHtmlOptions {
        /// <summary>
        /// Creates a readable semantic-document export with shared OfficeIMO document styling.
        /// </summary>
        public static WordToHtmlOptions CreateSemanticDocumentProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.WordLike) =>
            new WordToHtmlOptions {
                Profile = OfficeHtmlConversionProfile.WordSemanticDocument,
                Theme = theme,
                IncludeDefaultCss = true,
                UseSharedDocumentShell = true,
                IncludeFontStyles = true,
                IncludeListStyles = true,
                IncludeParagraphSpacingStyles = true,
                IncludeParagraphIndentationStyles = true,
                IncludeTableColumnGroups = true
            };

        /// <summary>
        /// Creates a trusted editable round-trip export with document structure and private review metadata enabled.
        /// </summary>
        public static WordToHtmlOptions CreateDocumentRoundTripProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.Report) =>
            new WordToHtmlOptions {
                Profile = OfficeHtmlConversionProfile.WordDocumentRoundTrip,
                Theme = theme,
                IncludeDefaultCss = true,
                UseSharedDocumentShell = true,
                IncludeFontStyles = true,
                IncludeListStyles = true,
                IncludeListDefinitions = true,
                IncludeParagraphClasses = true,
                IncludeRunClasses = true,
                IncludeParagraphSpacingStyles = true,
                IncludeParagraphIndentationStyles = true,
                ExportComments = true,
                ExportHeadersAndFooters = true,
                IncludeCustomProperties = true,
                IncludeSectionMetadata = true,
                IncludeTableColumnGroups = true
            };

        /// <summary>
        /// Creates a print-oriented review export with section, header, footer, and generated print styling.
        /// </summary>
        public static WordToHtmlOptions CreatePrintReviewProfile(OfficeVisualThemeKind theme = OfficeVisualThemeKind.WordLike) =>
            new WordToHtmlOptions {
                Profile = OfficeHtmlConversionProfile.WordPrintReview,
                Theme = theme,
                IncludeDefaultCss = true,
                UseSharedDocumentShell = true,
                IncludeFontStyles = true,
                IncludeListStyles = true,
                IncludeParagraphSpacingStyles = true,
                IncludeParagraphIndentationStyles = true,
                ExportComments = true,
                ExportHeadersAndFooters = true,
                IncludeCustomProperties = true,
                IncludeSectionMetadata = true,
                IncludeTableColumnGroups = true
            };

        /// <summary>Named Office-to-HTML fidelity contract represented by this export.</summary>
        public OfficeHtmlConversionProfile Profile { get; set; } = OfficeHtmlConversionProfile.WordSemanticDocument;

        /// <summary>Shared visual theme used when <see cref="IncludeDefaultCss"/> is enabled.</summary>
        public OfficeVisualThemeKind Theme { get; set; } = OfficeVisualThemeKind.WordLike;

        /// <summary>Maximum Open XML elements inspected during export. Defaults to 1,000,000.</summary>
        public long MaxDocumentElements { get; set; } = 1_000_000;

        /// <summary>Maximum bytes embedded for one image. Defaults to 64 MiB.</summary>
        public long MaxEmbeddedImageBytes { get; set; } = 64L * 1024 * 1024;

        /// <summary>Maximum aggregate image bytes embedded into one HTML result. Defaults to 256 MiB.</summary>
        public long MaxTotalEmbeddedImageBytes { get; set; } = 256L * 1024 * 1024;

        /// <summary>Maximum generated HTML characters. Defaults to 64 million.</summary>
        public long MaxOutputCharacters { get; set; } = 64_000_000;

        /// <summary>
        /// Maximum nested table depth exported from a Word document. The default is 128.
        /// </summary>
        public int MaxTableNestingDepth { get; set; } = 128;

        /// <summary>
        /// Maximum list nesting depth exported from a Word document. The default is 128.
        /// </summary>
        public int MaxListNestingDepth { get; set; } = 128;

        /// <summary>Maximum OMML equation nesting depth projected to text or MathML. Defaults to and cannot exceed 256.</summary>
        public int MaxEquationNestingDepth { get; set; } = 256;

        /// <summary>
        /// Optional font family applied to created runs during conversion.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>
        /// When true, includes run font information as inline styles.
        /// </summary>
        public bool IncludeFontStyles { get; set; }

        /// <summary>
        /// When set, includes list style information in generated HTML.
        /// </summary>
        public bool IncludeListStyles { get; set; }

        /// <summary>
        /// When true, emits reusable CSS classes and a head stylesheet for Word list definitions.
        /// Default is false to preserve the legacy inline/list-attribute output shape.
        /// </summary>
        public bool IncludeListDefinitions { get; set; }

        /// <summary>
        /// When true, paragraph styles are emitted as CSS classes.
        /// </summary>
        public bool IncludeParagraphClasses { get; set; }

        /// <summary>
        /// When true, run character styles are emitted as CSS classes.
        /// </summary>
        public bool IncludeRunClasses { get; set; }

        /// <summary>
        /// When true, includes run color information as inline styles. Defaults to true for conversion fidelity.
        /// </summary>
        public bool IncludeRunColorStyles { get; set; } = true;

        /// <summary>
        /// When true, includes run highlight information as inline styles. Defaults to true for conversion fidelity.
        /// </summary>
        public bool IncludeRunHighlightStyles { get; set; } = true;

        /// <summary>
        /// When true, includes paragraph spacing information as inline styles.
        /// </summary>
        public bool IncludeParagraphSpacingStyles { get; set; }

        /// <summary>
        /// When true, includes paragraph indentation information as inline styles.
        /// </summary>
        public bool IncludeParagraphIndentationStyles { get; set; }

        /// <summary>
        /// When true, footnotes are exported to HTML. Set to false to omit footnotes.
        /// </summary>
        public bool ExportFootnotes { get; set; } = true;

        /// <summary>
        /// When true, endnotes are exported to HTML. Set to false to omit endnotes.
        /// </summary>
        public bool ExportEndnotes { get; set; } = true;

        /// <summary>
        /// When true, Word comments are exported as linked HTML references and a comments section.
        /// Default is false so review metadata is not exposed unless requested.
        /// </summary>
        public bool ExportComments { get; set; }

        /// <summary>
        /// When true, Word section headers and footers are exported as semantic HTML
        /// <c>header</c> and <c>footer</c> regions with section/type metadata.
        /// Default is false to preserve the legacy body-only output.
        /// </summary>
        public bool ExportHeadersAndFooters { get; set; }

        /// <summary>
        /// When true, custom document properties are exported as typed HTML meta tags.
        /// Default is false so callers explicitly choose whether custom metadata is browser-visible.
        /// </summary>
        public bool IncludeCustomProperties { get; set; }

        /// <summary>
        /// When true, wraps exported document content in per-section <c>section</c>
        /// elements that preserve Word page size, orientation, and margin metadata.
        /// Default is false to preserve the legacy flat body output.
        /// </summary>
        public bool IncludeSectionMetadata { get; set; }

        /// <summary>
        /// When true, emits table column width metadata as HTML <c>colgroup</c>
        /// and <c>col</c> elements when the Word table exposes usable column widths.
        /// Default is false to preserve the legacy row-first table output.
        /// </summary>
        public bool IncludeTableColumnGroups { get; set; }

        /// <summary>
        /// When true (default), embeds images as base64 data URIs. When false,
        /// uses the image file paths instead.
        /// </summary>
        public bool EmbedImagesAsBase64 { get; set; } = true;

        /// <summary>
        /// Additional meta tags to include in the HTML head. Each tuple represents
        /// the <c>name</c> and <c>content</c> attributes of a meta element.
        /// </summary>
        public List<(string Name, string Content)> AdditionalMetaTags { get; } = new();

        /// <summary>
        /// Additional link tags to include in the HTML head. Each tuple represents
        /// the <c>rel</c> and <c>href</c> attributes of a link element.
        /// </summary>
        public List<(string Rel, string Href)> AdditionalLinkTags { get; } = new();

        /// <summary>
        /// When true, injects a small, built-in "Word-like" CSS into the HTML &lt;head&gt; to make output readable out-of-the-box.
        /// Default is false to preserve legacy behavior.
        /// </summary>
        public bool IncludeDefaultCss { get; set; } = false;

        /// <summary>
        /// Uses the shared responsive, print-aware OfficeIMO document shell when default CSS is included.
        /// Named export profiles enable this automatically. The default is false for legacy output compatibility.
        /// </summary>
        public bool UseSharedDocumentShell { get; set; }

        internal OfficeIMO.Html.HtmlDiagnosticReport ConversionReport { get; } = new OfficeIMO.Html.HtmlDiagnosticReport();

        internal WordToHtmlOptions CloneForConversion() {
            var clone = new WordToHtmlOptions {
                Profile = Profile,
                Theme = Theme,
                MaxDocumentElements = MaxDocumentElements,
                MaxEmbeddedImageBytes = MaxEmbeddedImageBytes,
                MaxTotalEmbeddedImageBytes = MaxTotalEmbeddedImageBytes,
                MaxOutputCharacters = MaxOutputCharacters,
                MaxTableNestingDepth = MaxTableNestingDepth,
                MaxListNestingDepth = MaxListNestingDepth,
                MaxEquationNestingDepth = MaxEquationNestingDepth,
                FontFamily = FontFamily,
                IncludeFontStyles = IncludeFontStyles,
                IncludeListStyles = IncludeListStyles,
                IncludeListDefinitions = IncludeListDefinitions,
                IncludeParagraphClasses = IncludeParagraphClasses,
                IncludeRunClasses = IncludeRunClasses,
                IncludeRunColorStyles = IncludeRunColorStyles,
                IncludeRunHighlightStyles = IncludeRunHighlightStyles,
                IncludeParagraphSpacingStyles = IncludeParagraphSpacingStyles,
                IncludeParagraphIndentationStyles = IncludeParagraphIndentationStyles,
                ExportFootnotes = ExportFootnotes,
                ExportEndnotes = ExportEndnotes,
                ExportComments = ExportComments,
                ExportHeadersAndFooters = ExportHeadersAndFooters,
                IncludeCustomProperties = IncludeCustomProperties,
                IncludeSectionMetadata = IncludeSectionMetadata,
                IncludeTableColumnGroups = IncludeTableColumnGroups,
                EmbedImagesAsBase64 = EmbedImagesAsBase64,
                IncludeDefaultCss = IncludeDefaultCss,
                UseSharedDocumentShell = UseSharedDocumentShell
            };
            clone.AdditionalMetaTags.AddRange(AdditionalMetaTags);
            clone.AdditionalLinkTags.AddRange(AdditionalLinkTags);
            return clone;
        }
    }
}
