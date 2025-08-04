using System;
using OfficeIMO.Word.Converters;

namespace OfficeIMO.Html {
    /// <summary>
    /// Options controlling Word to HTML conversion.
    /// </summary>
    public class WordToHtmlOptions : IConversionOptions {
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
