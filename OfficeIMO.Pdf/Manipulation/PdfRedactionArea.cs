namespace OfficeIMO.Pdf;

/// <summary>Rectangle requested for redaction planning, using PDF point coordinates from the page bottom-left.</summary>
public sealed class PdfRedactionArea {
    /// <summary>Creates a redaction area.</summary>
    public PdfRedactionArea(
        int pageNumber,
        double x,
        double y,
        double width,
        double height,
        string? label = null,
        PdfRedactionContentScope contentScope = PdfRedactionContentScope.TextAndUnderlay,
        PdfRedactionAppearanceMode appearanceMode = PdfRedactionAppearanceMode.Exact)
        : this(pageNumber, x, y, width, height, label, textRenderingMode: null, exactGeometry: null, contentScope, appearanceMode) {
    }

    private PdfRedactionArea(int pageNumber, double x, double y, double width, double height, string? label, int? textRenderingMode, PdfRedactionGeometry? exactGeometry, PdfRedactionContentScope contentScope, PdfRedactionAppearanceMode appearanceMode) {
        if (pageNumber < 1) {
            throw new ArgumentOutOfRangeException(nameof(pageNumber), "Page number must be greater than zero.");
        }

        if (!IsFinite(x)) {
            throw new ArgumentOutOfRangeException(nameof(x), "X coordinate must be finite.");
        }

        if (!IsFinite(y)) {
            throw new ArgumentOutOfRangeException(nameof(y), "Y coordinate must be finite.");
        }

        if (!IsFinite(width) || width <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(width), "Width must be finite and greater than zero.");
        }

        if (!IsFinite(height) || height <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(height), "Height must be finite and greater than zero.");
        }
        if (contentScope is < PdfRedactionContentScope.TextOnly or > PdfRedactionContentScope.TextAndUnderlay) throw new ArgumentOutOfRangeException(nameof(contentScope));
        if (appearanceMode is < PdfRedactionAppearanceMode.Exact or > PdfRedactionAppearanceMode.FullLine) throw new ArgumentOutOfRangeException(nameof(appearanceMode));

        PageNumber = pageNumber;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        Label = label;
        ContentScope = contentScope;
        AppearanceMode = appearanceMode;
        TextRenderingMode = textRenderingMode;
        ExactGeometry = exactGeometry;
    }

    /// <summary>One-based page number.</summary>
    public int PageNumber { get; }

    /// <summary>Left coordinate in PDF points.</summary>
    public double X { get; }

    /// <summary>Bottom coordinate in PDF points.</summary>
    public double Y { get; }

    /// <summary>Rectangle width in PDF points.</summary>
    public double Width { get; }

    /// <summary>Rectangle height in PDF points.</summary>
    public double Height { get; }

    /// <summary>Optional caller label.</summary>
    public string? Label { get; }

    /// <summary>Intersecting content removal policy for this reviewed area.</summary>
    public PdfRedactionContentScope ContentScope { get; }

    /// <summary>Visible privacy-appearance policy for this reviewed area.</summary>
    public PdfRedactionAppearanceMode AppearanceMode { get; }

    /// <summary>Right coordinate in PDF points.</summary>
    public double Right => X + Width;

    /// <summary>Top coordinate in PDF points.</summary>
    public double Top => Y + Height;

    internal int? TextRenderingMode { get; }

    internal PdfRedactionGeometry? ExactGeometry { get; }

    internal bool IntersectsRectangle(double x, double y, double width, double height) =>
        ExactGeometry?.IntersectsRectangle(x, y, width, height) ??
        X < x + width && Right > x && Y < y + height && Top > y;

    internal bool ContainsRectangle(double x, double y, double width, double height) =>
        ExactGeometry?.ContainsRectangle(x, y, width, height) ??
        x >= X && x + width <= Right && y >= Y && y + height <= Top;

    internal bool ContainsPoint(double x, double y) =>
        ExactGeometry?.ContainsPoint(x, y) ?? x >= X && x <= Right && y >= Y && y <= Top;

    internal bool IntersectsQuadrilateral(
        PdfRedactionPoint first,
        PdfRedactionPoint second,
        PdfRedactionPoint third,
        PdfRedactionPoint fourth) =>
        ExactGeometry?.IntersectsQuadrilateral(first, second, third, fourth) ??
        PdfRedactionGeometry.RectangleIntersectsQuadrilateral(X, Y, Width, Height, first, second, third, fourth);

    internal PdfRedactionArea WithExactGeometry(PdfRedactionGeometry exactGeometry) =>
        new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label, TextRenderingMode, exactGeometry, ContentScope, AppearanceMode);

    internal PdfRedactionArea WithTextRenderingMode(int textRenderingMode) =>
        new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label, textRenderingMode, ExactGeometry, ContentScope, AppearanceMode);

    internal PdfRedactionArea WithPolicies(PdfRedactionContentScope contentScope, PdfRedactionAppearanceMode appearanceMode) =>
        new PdfRedactionArea(PageNumber, X, Y, Width, Height, Label, TextRenderingMode, ExactGeometry, contentScope, appearanceMode);

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
