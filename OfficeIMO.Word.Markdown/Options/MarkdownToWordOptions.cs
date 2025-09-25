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
    }
}
