namespace OfficeIMO.Visio {
    /// <summary>
    /// Represents package-level Visio theme metadata preserved in a VSDX file.
    /// Authoring-time shape and connector styling is owned by <see cref="VisioStyleTheme"/>.
    /// </summary>
    public sealed class VisioPackageTheme {
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
