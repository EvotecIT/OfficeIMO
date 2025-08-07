namespace OfficeIMO.Word.Markdown {
    /// <summary>
    /// Specifies how images are exported when converting to Markdown.
    /// </summary>
    public enum ImageExportMode {
        /// <summary>
        /// Embed images directly into the Markdown as base64 data URIs.
        /// </summary>
        Base64,
        /// <summary>
        /// Save images to disk and reference them using relative file paths.
        /// </summary>
        File
    }
}
