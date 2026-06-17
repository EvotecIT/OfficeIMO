namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Controls how source package paths from OfficeIMO-native stencil catalog manifests are trusted.
    /// </summary>
    public sealed class VisioStencilCatalogManifestLoadOptions {
        /// <summary>
        /// Gets or sets whether SourcePackagePath values from the manifest should be retained.
        /// Defaults to false so untrusted manifests cannot trigger later local or UNC package reads.
        /// </summary>
        public bool AllowSourcePackagePaths { get; set; }

        /// <summary>
        /// Gets or sets whether retained SourcePackagePath values may point outside <see cref="BaseDirectory"/>.
        /// Defaults to false. Set this only for manifests and package paths from a trusted source.
        /// </summary>
        public bool AllowExternalSourcePackagePaths { get; set; }

        /// <summary>
        /// Gets or sets the directory used to resolve relative SourcePackagePath values.
        /// When loading from a file path, this defaults to the manifest file directory.
        /// </summary>
        public string? BaseDirectory { get; set; }
    }
}
