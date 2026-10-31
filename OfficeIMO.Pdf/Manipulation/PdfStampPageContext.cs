namespace OfficeIMO.Pdf;

/// <summary>Existing-page geometry supplied while building visual canvas stamp content.</summary>
public sealed class PdfStampPageContext {
    internal PdfStampPageContext(int pageNumber, int pageCount, double width, double height, int rotationDegrees) {
        PageNumber = pageNumber;
        PageCount = pageCount;
        Width = global::OfficeIMO.OfficeLength.Points(width);
        Height = global::OfficeIMO.OfficeLength.Points(height);
        RotationDegrees = rotationDegrees;
    }

    /// <summary>One-based target page number.</summary>
    public int PageNumber { get; }

    /// <summary>Total page count in the target PDF before stamping.</summary>
    public int PageCount { get; }

    /// <summary>Physical visual target-page width after crop, rotation, and UserUnit are applied.</summary>
    public global::OfficeIMO.OfficeLength Width { get; }

    /// <summary>Physical visual target-page height after crop, rotation, and UserUnit are applied.</summary>
    public global::OfficeIMO.OfficeLength Height { get; }

    /// <summary>Inherited target-page rotation normalized to 0, 90, 180, or 270 degrees.</summary>
    public int RotationDegrees { get; }
}
