namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one side of a table cell border.
/// </summary>
public sealed class PdfCellBorderSide {
    private global::OfficeIMO.OfficeLength _width = global::OfficeIMO.OfficeLength.Points(0.5);

    /// <summary>Border color. Set to null for no border on this side.</summary>
    public global::OfficeIMO.OfficeColor? Color { get; set; }

    /// <summary>Border stroke width as a physical length.</summary>
    public global::OfficeIMO.OfficeLength Width {
        get => _width;
        set {
            if (value.ToPoints() < 0 || double.IsNaN(value.ToPoints()) || double.IsInfinity(value.ToPoints())) {
                throw new System.ArgumentException("Table cell border widths must be non-negative finite values.", nameof(Width));
            }

            _width = value;
        }
    }

    /// <summary>Border stroke dash style.</summary>
    public OfficeIMO.Drawing.OfficeStrokeDashStyle DashStyle { get; set; }

    /// <summary>Border line style.</summary>
    public PdfCellBorderLineStyle LineStyle { get; set; }

    // Layout metadata belongs to the prepared table snapshot, not authored style.
    internal bool CenteredPair { get; set; }

    internal double PaintThickness => Color.HasValue && Width.ToPoints() > 0D
        ? Width.ToPoints() + (LineStyle == PdfCellBorderLineStyle.TwoLine ? DoubleTrackGap(Width.ToPoints()) : 0D) : 0D;

    internal static double DoubleTrackGap(double width) => Math.Max(width * 2D, 1D);

    /// <summary>Creates a copy of this table cell border side.</summary>
    public PdfCellBorderSide Clone() => new PdfCellBorderSide {
        Color = Color,
        Width = Width,
        DashStyle = DashStyle,
        LineStyle = LineStyle,
        CenteredPair = CenteredPair
    };
}
