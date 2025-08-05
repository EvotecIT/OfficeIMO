using System;
using OfficeIMO.Word;

namespace OfficeIMO.Html {
    /// <summary>
    /// Options controlling HTML to Word conversion.
    /// </summary>
    public class HtmlToWordOptions : ConversionOptions {
        /// <summary>
        /// When true, attempts to include list styling information during conversion.
        /// </summary>
        public bool IncludeListStyles { get; set; }
    }
}