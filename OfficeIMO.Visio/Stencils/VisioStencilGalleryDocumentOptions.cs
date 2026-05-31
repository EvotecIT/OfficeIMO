using OfficeIMO.Drawing;

namespace OfficeIMO.Visio.Stencils {
    /// <summary>
    /// Options for creating a complete Visio stencil gallery review document from a catalog.
    /// </summary>
    public sealed class VisioStencilGalleryDocumentOptions {
        /// <summary>
        /// Gets or sets the document title. When null, the catalog name is used.
        /// </summary>
        public string? Title { get; set; }

        /// <summary>
        /// Gets or sets the document author metadata.
        /// </summary>
        public string? Author { get; set; } = "OfficeIMO.Visio";

        /// <summary>
        /// Gets or sets whether generated gallery stencil instances should use document masters where possible.
        /// </summary>
        public bool UseMastersByDefault { get; set; } = true;

        /// <summary>
        /// Gets or sets whether an overview page with catalog counts and category summaries should be added.
        /// </summary>
        public bool IncludeOverviewPage { get; set; } = true;

        /// <summary>
        /// Gets or sets whether catalog pages should be grouped by stencil category.
        /// </summary>
        public bool GroupByCategory { get; set; } = true;

        /// <summary>
        /// Gets or sets the maximum number of stencil shapes rendered on each catalog page.
        /// </summary>
        public int ShapesPerPage { get; set; } = 24;

        /// <summary>
        /// Gets or sets the number of columns used on catalog pages.
        /// </summary>
        public int Columns { get; set; } = 4;

        /// <summary>
        /// Gets or sets the generated shape id prefix.
        /// </summary>
        public string IdPrefix { get; set; } = "stencil-gallery";

        /// <summary>
        /// Gets or sets the catalog page width.
        /// </summary>
        public double PageWidth { get; set; } = 11D;

        /// <summary>
        /// Gets or sets the catalog page height.
        /// </summary>
        public double PageHeight { get; set; } = 8.5D;

        /// <summary>
        /// Gets or sets the page measurement unit.
        /// </summary>
        public VisioMeasurementUnit PageUnit { get; set; } = VisioMeasurementUnit.Inches;

        /// <summary>
        /// Gets or sets whether catalog pages should grow when the selected grid needs more room.
        /// </summary>
        public bool AutoResizePages { get; set; } = true;

        /// <summary>
        /// Gets or sets whether placed stencil previews should receive visible Shape Data rows.
        /// </summary>
        public bool IncludeStencilMetadataShapeData { get; set; } = true;

        /// <summary>
        /// Gets or sets whether category labels should be shown in each gallery cell.
        /// </summary>
        public bool ShowCategory { get; set; } = true;

        /// <summary>
        /// Gets or sets the gallery page cell fill color.
        /// </summary>
        public OfficeColor CellFillColor { get; set; } = OfficeColor.FromRgb(248, 251, 254);

        /// <summary>
        /// Gets or sets the gallery page cell border color.
        /// </summary>
        public OfficeColor CellBorderColor { get; set; } = OfficeColor.FromRgb(209, 224, 239);

        /// <summary>
        /// Gets or sets the overview accent color.
        /// </summary>
        public OfficeColor AccentColor { get; set; } = OfficeColor.FromRgb(0, 120, 212);
    }
}
