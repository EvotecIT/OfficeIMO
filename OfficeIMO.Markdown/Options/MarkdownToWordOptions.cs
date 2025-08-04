using System;
using OfficeIMO.Word.Converters;

#nullable enable annotations

namespace OfficeIMO.Markdown {
    /// <summary>
    /// Options controlling Markdown to Word conversion.
    /// </summary>
    public class MarkdownToWordOptions : IConversionOptions {
        /// <summary>
        /// Optional font family applied to created runs.
        /// </summary>
        public string? FontFamily { get; set; }

        // Additional options may be added in the future.
    }
}
