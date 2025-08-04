using System;

namespace OfficeIMO.Markdown {
    /// <summary>
    /// Options controlling Markdown to Word conversion.
    /// </summary>
    public class MarkdownToWordOptions {
        /// <summary>
        /// Optional font family applied to created runs.
        /// </summary>
        public string? FontFamily { get; set; }

        // Additional options may be added in the future.
    }
}
