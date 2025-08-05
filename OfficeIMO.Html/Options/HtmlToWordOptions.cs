using System;
using OfficeIMO.Converters;

namespace OfficeIMO.Html {
    /// <summary>
    /// Options controlling HTML to Word conversion.
    /// </summary>
    public class HtmlToWordOptions : ConversionOptions {
        /// <summary>
        /// When true, attempts to keep list styling information during conversion.
        /// </summary>
        public bool PreserveListStyles { get; set; }
    }
}