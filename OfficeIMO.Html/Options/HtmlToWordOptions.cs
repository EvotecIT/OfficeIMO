using System;
using OfficeIMO.Word.Converters;

#nullable enable annotations

namespace OfficeIMO.Html {
    /// <summary>
    /// Options controlling HTML to Word conversion.
    /// </summary>
    public class HtmlToWordOptions : IConversionOptions {
        /// <summary>
        /// Optional font family applied to created runs.
        /// </summary>
        public string? FontFamily { get; set; }

        /// <summary>
        /// When true, attempts to keep list styling information during conversion.
        /// </summary>
        public bool PreserveListStyles { get; set; }
    }
}
