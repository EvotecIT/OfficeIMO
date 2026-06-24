namespace OfficeIMO.Excel {
    /// <summary>
    /// Describes the workbook theme package part.
    /// </summary>
    public sealed class ExcelWorkbookThemeInfo {
        /// <summary>
        /// Creates workbook theme information.
        /// </summary>
        public ExcelWorkbookThemeInfo(bool hasTheme, string? name, string? xml) {
            HasTheme = hasTheme;
            Name = name;
            Xml = xml;
        }

        /// <summary>
        /// Gets whether the workbook contains a theme part.
        /// </summary>
        public bool HasTheme { get; }

        /// <summary>
        /// Gets the theme name from the theme XML root when available.
        /// </summary>
        public string? Name { get; }

        /// <summary>
        /// Gets the theme XML when requested by callers.
        /// </summary>
        public string? Xml { get; }
    }
}
