using System;
using System.Collections.Generic;
using OfficeIMO.Word;

namespace OfficeIMO.Word.Html {
    /// <summary>
    /// Options controlling Word to HTML conversion.
    /// </summary>
    public class WordToHtmlOptions {
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
        /// When true, paragraph styles are emitted as CSS classes.
        /// </summary>
        public bool IncludeParagraphClasses { get; set; }

        /// <summary>
        /// When true, run character styles are emitted as CSS classes.
        /// </summary>
        public bool IncludeRunClasses { get; set; }

        /// <summary>
        /// When true, footnotes are exported to HTML. Set to false to omit footnotes.
        /// </summary>
        public bool ExportFootnotes { get; set; } = true;

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
    }
}
