namespace OfficeIMO.Pdf;

/// <summary>Commonly used page sizes.</summary>
public static class PageSizes {
    /// <summary>ISO A5 size (420 × 595 pt).</summary>
    public static PageSize A5 => new PageSize(420, 595);
    /// <summary>ISO A4 size (595 × 842 pt).</summary>
    public static PageSize A4 => new PageSize(595, 842);
    /// <summary>US Letter size (612 × 792 pt).</summary>
    public static PageSize Letter => new PageSize(612, 792);
    /// <summary>US Legal size (612 × 1008 pt).</summary>
    public static PageSize Legal => new PageSize(612, 1008);
}
