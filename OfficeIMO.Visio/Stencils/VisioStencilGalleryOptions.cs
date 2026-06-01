using OfficeIMO.Drawing;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Layout and visual options for rendering a stencil catalog as a browsable Visio page gallery.
    /// </summary>
    public sealed class VisioStencilGalleryOptions {
        /// <summary>
        /// Gets or sets the gallery title. When null, the catalog name is used.
        /// </summary>
        public string? Title { get; set; }

        /// <summary>
        /// Gets or sets the shape id prefix used for generated gallery elements.
        /// </summary>
        public string IdPrefix { get; set; } = "stencil-gallery";

        /// <summary>
        /// Gets or sets the maximum number of shapes to render.
        /// </summary>
        public int MaxShapes { get; set; } = 24;

        /// <summary>
        /// Gets or sets the number of columns in the gallery grid.
        /// </summary>
        public int Columns { get; set; } = 4;

        /// <summary>
        /// Gets or sets the left page margin in inches.
        /// </summary>
        public double Left { get; set; } = 0.65D;

        /// <summary>
        /// Gets or sets the top page margin in inches.
        /// </summary>
        public double Top { get; set; } = 0.55D;

        /// <summary>
        /// Gets or sets the horizontal gap between gallery cells in inches.
        /// </summary>
        public double ColumnGap { get; set; } = 0.25D;

        /// <summary>
        /// Gets or sets the vertical gap between gallery cells in inches.
        /// </summary>
        public double RowGap { get; set; } = 0.25D;

        /// <summary>
        /// Gets or sets the width of each gallery cell in inches.
        /// </summary>
        public double CellWidth { get; set; } = 2.15D;

        /// <summary>
        /// Gets or sets the height of each gallery cell in inches.
        /// </summary>
        public double CellHeight { get; set; } = 1.42D;

        /// <summary>
        /// Gets or sets the maximum icon width inside a gallery cell in inches.
        /// </summary>
        public double IconMaxWidth { get; set; } = 0.9D;

        /// <summary>
        /// Gets or sets the maximum icon height inside a gallery cell in inches.
        /// </summary>
        public double IconMaxHeight { get; set; } = 0.72D;

        /// <summary>
        /// Gets or sets whether the page should grow to fit the requested gallery grid.
        /// </summary>
        public bool AutoResizePage { get; set; } = true;

        /// <summary>
        /// Gets or sets whether category labels should be shown in each gallery cell.
        /// </summary>
        public bool ShowCategory { get; set; } = true;

        /// <summary>
        /// Gets or sets whether each placed stencil preview should receive visible Shape Data rows for catalog review.
        /// </summary>
        public bool IncludeStencilMetadataShapeData { get; set; }

        /// <summary>
        /// Gets or sets the title text color.
        /// </summary>
        public OfficeColor TitleColor { get; set; } = OfficeColor.FromRgb(31, 48, 63);

        /// <summary>
        /// Gets or sets the gallery cell fill color.
        /// </summary>
        public OfficeColor CellFillColor { get; set; } = OfficeColor.FromRgb(248, 251, 254);

        /// <summary>
        /// Gets or sets the gallery cell border color.
        /// </summary>
        public OfficeColor CellBorderColor { get; set; } = OfficeColor.FromRgb(209, 224, 239);
    }
}
