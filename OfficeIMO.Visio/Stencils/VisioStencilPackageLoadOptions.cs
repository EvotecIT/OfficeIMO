using System.Collections.Generic;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Options for creating an OfficeIMO-native stencil catalog from a Visio package.
    /// </summary>
    public sealed class VisioStencilPackageLoadOptions {
        /// <summary>
        /// Gets or sets the catalog display name. Defaults to the package file name.
        /// </summary>
        public string? CatalogName { get; set; }

        /// <summary>
        /// Gets or sets the category assigned to loaded stencil shapes. Defaults to the catalog name.
        /// </summary>
        public string? Category { get; set; }

        /// <summary>
        /// Gets or sets a stable id prefix. Defaults to the normalized package file name.
        /// </summary>
        public string? IdPrefix { get; set; }

        /// <summary>
        /// Gets or sets optional master NameU filters.
        /// </summary>
        public IEnumerable<string>? MasterNames { get; set; }

        /// <summary>
        /// Gets or sets whether unsupported masters should be included as generic generated masters.
        /// Defaults to false so package loading remains a learning/catalog step, not template ingestion.
        /// </summary>
        public bool IncludeUnsupportedMasters { get; set; }

        /// <summary>
        /// Gets or sets the default stencil width in the caller's placement unit.
        /// </summary>
        public double DefaultWidth { get; set; } = 1.8;

        /// <summary>
        /// Gets or sets the default stencil height in the caller's placement unit.
        /// </summary>
        public double DefaultHeight { get; set; } = 0.9;
    }
}
