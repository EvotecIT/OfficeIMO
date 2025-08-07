using System;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Options controlling Word to HTML conversion.
    /// </summary>
    public class WordToHtmlOptions {
        /// <summary>
        /// Optional font family applied to created runs during conversion.
        /// </summary>
        public string FontFamily { get; set; }
        
        /// <summary>
        /// When true, includes run font information as inline styles.
        /// </summary>
        public bool IncludeFontStyles { get; set; }

        /// <summary>
        /// When set, includes list style information in generated HTML.
        /// </summary>
        public bool IncludeListStyles { get; set; }

        /// <summary>
        /// When true, footnotes are exported to HTML. Set to false to omit footnotes.
        /// </summary>
        public bool ExportFootnotes { get; set; } = true;

        /// <summary>
        /// When true (default), embeds images as base64 data URIs. When false,
        /// uses the image file paths instead.
        /// </summary>
        public bool EmbedImagesAsBase64 { get; set; } = true;
    }
}