namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one side of a panel border.
/// </summary>
public sealed class PdfPanelBorder {
    private double _width = 0.5;

    /// <summary>Border color. Set to null for no border on this side.</summary>
    public PdfColor? Color { get; set; }

    /// <summary>Border stroke width in points.</summary>
    public double Width {
        get => _width;
        set {
            if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentException("Panel border width must be a non-negative finite value.", nameof(Width));
            }

            _width = value;
        }
    }

    /// <summary>Creates a copy of this panel border.</summary>
    public PdfPanelBorder Clone() => new PdfPanelBorder {
        Color = Color,
        Width = Width
    };
}
