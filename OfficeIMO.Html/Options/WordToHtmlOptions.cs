using System;
using OfficeIMO.Converters;

namespace OfficeIMO.Html {
    /// <summary>
    /// Options controlling Word to HTML conversion.
    /// </summary>
    public class WordToHtmlOptions : ConversionOptions {
        /// <summary>
        /// When true, includes run font information as inline styles.
        /// </summary>
        public bool IncludeStyles { get; set; }

        /// <summary>
        /// When set, retains list style information in generated HTML.
        /// </summary>
        public bool PreserveListStyles { get; set; }
    }
}