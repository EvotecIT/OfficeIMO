namespace OfficeIMO.Pdf;

/// <summary>Represents a physical PDF page size.</summary>
public readonly struct PdfPageSize {
    /// <summary>Page width.</summary>
    public OfficeLength Width { get; }
    /// <summary>Page height.</summary>
    public OfficeLength Height { get; }
    /// <summary>Creates a page size with explicit units.</summary>
    public PdfPageSize(OfficeLength width, OfficeLength height) {
        if (width.ToEmus() <= 0) throw new ArgumentOutOfRangeException(nameof(width));
        if (height.ToEmus() <= 0) throw new ArgumentOutOfRangeException(nameof(height));
        Width = width;
        Height = height;
    }

    /// <summary>Returns this size in portrait orientation.</summary>
    public PdfPageSize Portrait() => Width.CompareTo(Height) <= 0 ? this : new PdfPageSize(Height, Width);

    /// <summary>Returns this size in landscape orientation.</summary>
    public PdfPageSize Landscape() => Width.CompareTo(Height) >= 0 ? this : new PdfPageSize(Height, Width);

    /// <summary>Returns this size in the requested orientation.</summary>
    public PdfPageSize WithOrientation(OfficePageOrientation orientation) {
        Guard.PageOrientation(orientation, nameof(orientation));
        return orientation == OfficePageOrientation.Landscape ? Landscape() : Portrait();
    }
}
