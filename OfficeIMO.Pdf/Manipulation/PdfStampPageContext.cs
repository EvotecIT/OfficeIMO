namespace OfficeIMO.Pdf;

/// <summary>Existing-page geometry supplied while building visual canvas stamp content.</summary>
public sealed class PdfStampPageContext {
    internal PdfStampPageContext(int pageNumber, int pageCount, double width, double height, int rotationDegrees) {
        PageNumber = pageNumber;
        PageCount = pageCount;
        Width = width;
        Height = height;
        RotationDegrees = rotationDegrees;
    }

    /// <summary>One-based target page number.</summary>
    public int PageNumber { get; }

    /// <summary>Total page count in the target PDF before stamping.</summary>
    public int PageCount { get; }

    /// <summary>Visual target-page width in points after crop and rotation are applied.</summary>
    public double Width { get; }

    /// <summary>Visual target-page height in points after crop and rotation are applied.</summary>
    public double Height { get; }

    /// <summary>Inherited target-page rotation normalized to 0, 90, 180, or 270 degrees.</summary>
    public int RotationDegrees { get; }
}
