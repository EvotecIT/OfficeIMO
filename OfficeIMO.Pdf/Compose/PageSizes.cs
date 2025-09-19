namespace OfficeIMO.Pdf;

/// <summary>Commonly used page sizes.</summary>
public static class PageSizes {
    public static PageSize A5 => new PageSize(420, 595);
    public static PageSize A4 => new PageSize(595, 842);
    public static PageSize Letter => new PageSize(612, 792);
    public static PageSize Legal => new PageSize(612, 1008);
}

