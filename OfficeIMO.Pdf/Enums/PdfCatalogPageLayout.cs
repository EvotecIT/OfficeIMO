namespace OfficeIMO.Pdf;

/// <summary>
/// Catalog page layout requested for the initial page arrangement of a generated PDF.
/// </summary>
public enum PdfCatalogPageLayout {
    /// <summary>
    /// Displays one page at a time.
    /// </summary>
    SinglePage,

    /// <summary>
    /// Displays pages in one continuous column.
    /// </summary>
    OneColumn,

    /// <summary>
    /// Displays pages in two columns with odd-numbered pages on the left.
    /// </summary>
    TwoColumnLeft,

    /// <summary>
    /// Displays pages in two columns with odd-numbered pages on the right.
    /// </summary>
    TwoColumnRight,

    /// <summary>
    /// Displays pages two at a time with odd-numbered pages on the left.
    /// </summary>
    TwoPageLeft,

    /// <summary>
    /// Displays pages two at a time with odd-numbered pages on the right.
    /// </summary>
    TwoPageRight
}
