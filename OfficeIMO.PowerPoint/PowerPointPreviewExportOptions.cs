namespace OfficeIMO.PowerPoint {
    /// <summary>
    ///     Options for the optional Microsoft PowerPoint automation preview/export boundary.
    /// </summary>
    public sealed class PowerPointPreviewExportOptions {
        /// <summary>
        ///     Gets or sets whether the caller has classified the presentation as trusted for
        ///     opening in installed Microsoft PowerPoint automation. Keep this disabled for
        ///     user-supplied or otherwise untrusted presentation files.
        /// </summary>
        public bool TrustPresentationFile { get; set; }
    }
}
