using System;
using OfficeIMO.Word;

namespace OfficeIMO.Html {
    /// <summary>
    /// Options controlling Word to HTML conversion.
    /// </summary>
    public class WordToHtmlOptions : ConversionOptions {
        /// <summary>
        /// When true, includes run font information as inline styles.
        /// </summary>
        public bool IncludeFontStyles { get; set; }

        /// <summary>
        /// When set, includes list style information in generated HTML.
        /// </summary>
        public bool IncludeListStyles { get; set; }
    }
}