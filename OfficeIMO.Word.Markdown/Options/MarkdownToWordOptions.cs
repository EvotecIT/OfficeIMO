using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Options controlling Markdown to Word conversion.
    /// </summary>
    public partial class WordMarkdownImportOptions {
        private bool? _renderFrontMatter;
        internal WordMarkdownImportOptions CloneForConversion() {
            var copy = (WordMarkdownImportOptions)MemberwiseClone();
            copy._theme = _theme?.Clone();
            copy.ReaderOptions = ReaderOptions?.Clone();
            copy.AllowedImageDirectories = new System.Collections.Generic.HashSet<string>(AllowedImageDirectories, AllowedImageDirectories.Comparer);
            copy.AllowedImageSchemes = new System.Collections.Generic.HashSet<string>(AllowedImageSchemes, AllowedImageSchemes.Comparer);
            copy.ImageLayout = ImageLayout.CloneForConversion();
            return copy;
        }

        /// <summary>Rejects fidelity loss before returning the target document or writing its destination.</summary>
        public bool FailOnLoss { get; set; }
        /// <summary>Optional caller-owned template to populate rather than creating a detached document.</summary>
        public WordDocument? Template { get; set; }
        /// <summary>Tag identifying the template insertion content control.</summary>
        public string? ContentControlTag { get; set; }
        /// <summary>Alias identifying the template insertion content control.</summary>
        public string? ContentControlAlias { get; set; }
        /// <summary>Bookmark identifying the insertion point.</summary>
        public string? BookmarkName { get; set; }
        /// <summary>Replaces the selected placeholder after insertion.</summary>
        public bool ReplacePlaceholder { get; set; } = true;

        internal Func<string, WordDocument, Action<WordMarkdownConversionDiagnostic>?, bool>? RawHtmlRenderer { get; set; }
        internal Func<string, Action<WordMarkdownConversionDiagnostic>?, OfficeIMO.Markdown.MarkdownDocument?>? RawHtmlParser { get; set; }
        internal Action<WordMarkdownConversionDiagnostic>? DiagnosticSink { get; set; }

        private OfficeIMO.Markdown.MarkdownVisualTheme? _theme;

        /// <summary>
        /// Shared Markdown visual theme used to align generated Word output with Markdown HTML and PDF exports.
        /// </summary>
        public OfficeIMO.Markdown.MarkdownVisualTheme? Theme {
            get => _theme?.Clone();
            set => _theme = value?.Clone();
        }

        internal OfficeIMO.Markdown.MarkdownVisualTheme? ThemeSnapshot => _theme?.Clone();

        /// <summary>
        /// Applies the shared default visual theme when <see cref="Theme"/> is omitted.
        /// Set to <see langword="false"/> for plain output that relies only on Word defaults.
        /// </summary>
        public bool ApplyDefaultTheme { get; set; } = true;

        /// <summary>
        /// Optional font family applied to created runs during conversion.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>
        /// Optional base URI applied when resolving relative links and images in Markdown.
        /// </summary>
        public string? BaseUri { get; set; }

        /// <summary>
        /// Emits warnings from the converter (e.g., invalid URIs). Optional.
        /// </summary>
        public Action<string>? OnWarning { get; set; }

        /// <summary>
        /// Maximum nested block-rendering depth accepted by the Word converter. Default: 64.
        /// </summary>
        public int MaxNestingDepth { get; set; } = 64;

        /// <summary>
        /// Optional default page size applied when creating new documents.
        /// </summary>
        public WordPageSize? DefaultPageSize { get; set; }

        /// <summary>
        /// Optional default page orientation applied when creating new documents.
        /// </summary>
        public OfficePageOrientation? DefaultOrientation { get; set; }

        /// <summary>
        /// Applies default page settings to the provided document instance.
        /// </summary>
        internal void ApplyDefaults(WordDocument document) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (DefaultPageSize.HasValue) {
                document.PageSettings.PageSize = DefaultPageSize.Value;
            }

            if (DefaultOrientation.HasValue) {
                document.PageOrientation = DefaultOrientation.Value;
            }
        }

        // Image handling security knobs
        /// <summary>Allow inserting local images from file paths. Default: false.</summary>
        public bool AllowLocalImages { get; set; }
        /// <summary>Restrict local images to these directories (optional). Paths are normalized for comparison.</summary>
        public System.Collections.Generic.HashSet<string> AllowedImageDirectories { get; private set; } = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
        /// <summary>
        /// Optional resolver for remote image bytes. The converter never performs network requests itself;
        /// callers own download, cache, authentication, and trust policy.
        /// </summary>
        public Func<Uri, byte[]?>? RemoteImageResolver { get; set; }
        /// <summary>
        /// Allow embedding base64 <c>data:image/*</c> Markdown images. Default: true so the
        /// default Word-to-Markdown image export mode can round-trip back into Word.
        /// </summary>
        public bool AllowDataUriImages { get; set; } = true;
        /// <summary>
        /// Maximum decoded size for a single base64 data URI image. Default: 32 MiB.
        /// </summary>
        public long MaxDataUriImageBytes { get; set; } = 32L * 1024L * 1024L;
        /// <summary>Allowed URL schemes for remote images. Default: http, https.</summary>
        public System.Collections.Generic.HashSet<string> AllowedImageSchemes { get; private set; } = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase) { "http", "https" };
        /// <summary>Optional custom validator for remote image URLs.</summary>
        public Func<System.Uri, bool>? ImageUrlValidator { get; set; }
        /// <summary>When remote images are not allowed or validation fails, insert a hyperlink instead of image. Default: true.</summary>
        public bool FallbackRemoteImagesToHyperlinks { get; set; } = true;
        /// <summary>Maximum byte length accepted from <see cref="RemoteImageResolver"/> for one image.</summary>
        public long MaximumRemoteImageBytes { get; set; } = 5L * 1024L * 1024L;

        /// <summary>
        /// Optional callback receiving per-image layout diagnostics.
        /// </summary>
        public Action<WordMarkdownImageLayoutDiagnostic>? OnImageLayoutDiagnostic { get; set; }

        /// <summary>
        /// Image sizing and fitting behavior for markdown conversion.
        /// </summary>
        public WordMarkdownImageLayoutOptions ImageLayout { get; private set; } = new WordMarkdownImageLayoutOptions();

        /// <summary>
        /// When enabled, isolated single-line <c>Label: value</c> content is kept as
        /// narrative paragraphs while grouped definition-list blocks remain structured.
        /// </summary>
        public bool PreferNarrativeSingleLineDefinitions { get; set; }

        /// <summary>
        /// Optional markdown reader configuration used before Word conversion.
        /// When omitted, the converter uses the default OfficeIMO reader profile.
        /// </summary>
        public OfficeIMO.Markdown.MarkdownLoadOptions? ReaderOptions { get; set; }

        /// <summary>
        /// Creates the effective Markdown reader options for a Word import, including the
        /// Word semantic fenced-block extensions and this instance's base URI settings.
        /// </summary>
        public OfficeIMO.Markdown.MarkdownLoadOptions CreateReaderOptions() {
            OfficeIMO.Markdown.MarkdownLoadOptions effective;
            if (ReaderOptions == null) {
                effective = new OfficeIMO.Markdown.MarkdownLoadOptions {
                    BaseUri = BaseUri,
                    PreferNarrativeSingleLineDefinitions = PreferNarrativeSingleLineDefinitions
                };
            } else {
                effective = ReaderOptions.Clone();
                if (string.IsNullOrWhiteSpace(effective.BaseUri) && !string.IsNullOrWhiteSpace(BaseUri)) {
                    effective.BaseUri = BaseUri;
                }
                if (!effective.PreferNarrativeSingleLineDefinitions && PreferNarrativeSingleLineDefinitions) {
                    effective.PreferNarrativeSingleLineDefinitions = true;
                }
            }

            WordMarkdownSemanticBlocks.ConfigureReaderOptions(effective);
            return effective;
        }

        /// <summary>
        /// Controls whether front matter is rendered as visible Word content.
        /// The default is true for a new document and false when a template consumes it as metadata.
        /// </summary>
        public bool RenderFrontMatter {
            get => _renderFrontMatter ?? Template == null;
            set => _renderFrontMatter = value;
        }

        /// <summary>
        /// Optional hard cap (pixels) applied to rendered markdown image width.
        /// </summary>
        public double? MaxImageWidthPixels {
            get => ImageLayout.MaxWidthPixels;
            set => ImageLayout.MaxWidthPixels = value;
        }

        /// <summary>
        /// Optional hard cap (pixels) applied to rendered markdown image height.
        /// </summary>
        public double? MaxImageHeightPixels {
            get => ImageLayout.MaxHeightPixels;
            set => ImageLayout.MaxHeightPixels = value;
        }

        /// <summary>
        /// Optional hard cap expressed as percent of available content width (for example 100, 85, 50).
        /// </summary>
        public double? MaxImageWidthPercentOfContent {
            get => ImageLayout.MaxWidthPercentOfContent;
            set => ImageLayout.MaxWidthPercentOfContent = value;
        }

        /// <summary>
        /// When enabled, markdown images are constrained to section content width
        /// (page width minus left/right margins).
        /// </summary>
        public bool FitImagesToPageContentWidth {
            get => ImageLayout.FitMode == WordMarkdownImageFitMode.PageContentWidth;
            set {
                if (value) {
                    ImageLayout.FitMode = WordMarkdownImageFitMode.PageContentWidth;
                } else if (ImageLayout.FitMode == WordMarkdownImageFitMode.PageContentWidth) {
                    ImageLayout.FitMode = WordMarkdownImageFitMode.None;
                }
            }
        }

        /// <summary>
        /// When enabled, markdown images are constrained to context width
        /// (section content width minus list/quote indentation).
        /// </summary>
        public bool FitImagesToContextWidth {
            get => ImageLayout.FitMode == WordMarkdownImageFitMode.ContextContentWidth;
            set {
                if (value) {
                    ImageLayout.FitMode = WordMarkdownImageFitMode.ContextContentWidth;
                } else if (ImageLayout.FitMode == WordMarkdownImageFitMode.ContextContentWidth) {
                    ImageLayout.FitMode = WordMarkdownImageFitMode.None;
                }
            }
        }
    }
}
