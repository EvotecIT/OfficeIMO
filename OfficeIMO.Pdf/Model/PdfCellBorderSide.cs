namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one side of a table cell border.
/// </summary>
public sealed class PdfCellBorderSide {
    private double _width = 0.5;

    /// <summary>Border color. Set to null for no border on this side.</summary>
    public PdfColor? Color { get; set; }

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

    /// <summary>Border stroke dash style.</summary>
    public OfficeIMO.Drawing.OfficeStrokeDashStyle DashStyle { get; set; }

    /// <summary>Border line style.</summary>
    public PdfCellBorderLineStyle LineStyle { get; set; }

    /// <summary>Creates a copy of this table cell border side.</summary>
    public PdfCellBorderSide Clone() => new PdfCellBorderSide {
        Color = Color,
        Width = Width,
        DashStyle = DashStyle,
        LineStyle = LineStyle
    };
}
