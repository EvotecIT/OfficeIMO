using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Options controlling Markdown to Word conversion.
    /// </summary>
    public class MarkdownToWordOptions {
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
        /// Optional default page size applied when creating new documents.
        /// </summary>
        public WordPageSize? DefaultPageSize { get; set; }

        /// <summary>
        /// Optional default page orientation applied when creating new documents.
        /// </summary>
        public PageOrientationValues? DefaultOrientation { get; set; }

        /// <summary>
        /// Applies default page settings to the provided document instance.
        /// </summary>
        public void ApplyDefaults(WordDocument document) {
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
        public System.Collections.Generic.HashSet<string> AllowedImageDirectories { get; } = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase);
        /// <summary>Allow downloading remote images. Default: false.</summary>
        public bool AllowRemoteImages { get; set; }
        /// <summary>Allowed URL schemes for remote images. Default: http, https.</summary>
        public System.Collections.Generic.HashSet<string> AllowedImageSchemes { get; } = new System.Collections.Generic.HashSet<string>(System.StringComparer.OrdinalIgnoreCase) { "http", "https" };
        /// <summary>Optional custom validator for remote image URLs.</summary>
        public Func<System.Uri, bool>? ImageUrlValidator { get; set; }
        /// <summary>When remote images are not allowed or validation fails, insert a hyperlink instead of image. Default: true.</summary>
        public bool FallbackRemoteImagesToHyperlinks { get; set; } = true;
        /// <summary>Timeout applied to remote image downloads. Default: 20 seconds.</summary>
        public TimeSpan RemoteImageDownloadTimeout { get; set; } = TimeSpan.FromSeconds(20);

        /// <summary>
        /// Optional callback receiving per-image layout diagnostics.
        /// </summary>
        public Action<MarkdownImageLayoutDiagnostic>? OnImageLayoutDiagnostic { get; set; }

        /// <summary>
        /// Image sizing and fitting behavior for markdown conversion.
        /// </summary>
        public MarkdownImageLayoutOptions ImageLayout { get; } = new MarkdownImageLayoutOptions();

        /// <summary>
        /// When enabled, markdown parser definition-list detection is disabled so
        /// <c>Label: value</c> lines are kept as narrative paragraphs.
        /// </summary>
        public bool PreferNarrativeSingleLineDefinitions { get; set; }

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
        /// When enabled, SVG sources are rasterized to PNG before insertion.
        /// </summary>
        public bool PreferRasterizeSvgForWord {
            get => ImageLayout.PreferRasterizeSvgForWord;
            set => ImageLayout.PreferRasterizeSvgForWord = value;
        }

        /// <summary>
        /// Rasterization DPI applied to SVG sources when rasterization is enabled.
        /// </summary>
        public int SvgRasterizationDpi {
            get => ImageLayout.SvgRasterizationDpi;
            set => ImageLayout.SvgRasterizationDpi = value;
        }

        /// <summary>
        /// When enabled, markdown images are constrained to section content width
        /// (page width minus left/right margins).
        /// </summary>
        public bool FitImagesToPageContentWidth {
            get => ImageLayout.FitMode == MarkdownImageFitMode.PageContentWidth;
            set {
                if (value) {
                    ImageLayout.FitMode = MarkdownImageFitMode.PageContentWidth;
                } else if (ImageLayout.FitMode == MarkdownImageFitMode.PageContentWidth) {
                    ImageLayout.FitMode = MarkdownImageFitMode.None;
                }
            }
        }

        /// <summary>
        /// When enabled, markdown images are constrained to context width
        /// (section content width minus list/quote indentation).
        /// </summary>
        public bool FitImagesToContextWidth {
            get => ImageLayout.FitMode == MarkdownImageFitMode.ContextContentWidth;
            set {
                if (value) {
                    ImageLayout.FitMode = MarkdownImageFitMode.ContextContentWidth;
                } else if (ImageLayout.FitMode == MarkdownImageFitMode.ContextContentWidth) {
                    ImageLayout.FitMode = MarkdownImageFitMode.None;
                }
            }
        }
    }
}
