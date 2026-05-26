namespace OfficeIMO.Pdf;

/// <summary>
/// Describes an override border for a single table cell.
/// </summary>
public sealed class PdfCellBorder {
    private double _width = 0.5;

    /// <summary>Border color. Set to null or use a zero width to suppress the border.</summary>
    public PdfColor? Color { get; set; } = new PdfColor(0.8, 0.8, 0.8);

    /// <summary>Border stroke width in points.</summary>
    public double Width {
        get => _width;
        set {
            if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentException("Table cell border widths must be non-negative finite values.", nameof(Width));
            }

            _width = value;
        }
    }

    /// <summary>Whether to draw the top side of the cell border.</summary>
    public bool Top { get; set; } = true;

    /// <summary>Whether to draw the right side of the cell border.</summary>
    public bool Right { get; set; } = true;

    /// <summary>Whether to draw the bottom side of the cell border.</summary>
    public bool Bottom { get; set; } = true;

    /// <summary>Whether to draw the left side of the cell border.</summary>
    public bool Left { get; set; } = true;

    /// <summary>Creates a deep copy of this border style.</summary>
    public PdfCellBorder Clone() => new PdfCellBorder {
        Color = Color,
        Width = Width,
        Top = Top,
        Right = Right,
        Bottom = Bottom,
        Left = Left
    };
}
