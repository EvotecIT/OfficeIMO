namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents a Visio theme.
    /// </summary>
    public class VisioTheme {
        /// <summary>
        /// Name of the theme.
        /// </summary>
        public string? Name { get; set; }

        /// <summary>
        /// Raw theme XML captured from an existing package so it can be preserved on save.
        /// </summary>
        internal System.Xml.Linq.XDocument? TemplateXml { get; set; }
    }
}
