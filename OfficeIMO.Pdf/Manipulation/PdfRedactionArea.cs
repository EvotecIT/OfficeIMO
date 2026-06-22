namespace OfficeIMO.Pdf;

/// <summary>Rectangle requested for redaction planning, using PDF point coordinates from the page bottom-left.</summary>
public sealed class PdfRedactionArea {
    /// <summary>Creates a redaction area.</summary>
    public PdfRedactionArea(int pageNumber, double x, double y, double width, double height, string? label = null) {
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

        PageNumber = pageNumber;
        X = x;
        Y = y;
        Width = width;
        Height = height;
        Label = label;
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

    /// <summary>Right coordinate in PDF points.</summary>
    public double Right => X + Width;

    /// <summary>Top coordinate in PDF points.</summary>
    public double Top => Y + Height;

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
}
