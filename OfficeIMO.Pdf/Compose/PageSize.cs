namespace OfficeIMO.Pdf;

/// <summary>Represents a page size in points.</summary>
public readonly struct PageSize {
    /// <summary>Width in points.</summary>
    public double Width { get; }
    /// <summary>Height in points.</summary>
    public double Height { get; }
    /// <summary>Creates a new page size.</summary>
    public PageSize(double width, double height) { Width = width; Height = height; }
}

