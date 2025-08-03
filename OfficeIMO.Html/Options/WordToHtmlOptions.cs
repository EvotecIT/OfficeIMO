using System;

namespace OfficeIMO.Html {
    /// <summary>
    /// Options controlling Word to HTML conversion.
    /// </summary>
    public class WordToHtmlOptions {
        /// <summary>
        /// When true, includes run font information as inline styles.
        /// </summary>
        public bool IncludeStyles { get; set; }
    }
}
