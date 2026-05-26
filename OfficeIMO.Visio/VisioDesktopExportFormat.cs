namespace OfficeIMO.Visio {
    /// <summary>
    /// Export formats supported by <see cref="VisioDesktopValidator"/> when Microsoft Visio desktop is available.
    /// </summary>
    public enum VisioDesktopExportFormat {
        /// <summary>
        /// Scalable Vector Graphics export of the first page.
        /// </summary>
        Svg,

        /// <summary>
        /// PNG image export of the first page.
        /// </summary>
        Png,

        /// <summary>
        /// PDF export of the document.
        /// </summary>
        Pdf
    }
}
