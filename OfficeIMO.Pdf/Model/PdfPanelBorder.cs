namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one side of a panel border.
/// </summary>
public sealed class PdfPanelBorder {
    private double _width = 0.5;
    private double _offset;

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

    /// <summary>
    /// Distance in points from the panel boundary to this border's stroke center.
    /// Positive values move the stroke outward; negative values move it inward.
    /// This affects painting without changing the panel's content width or padding.
    /// </summary>
    public double Offset {
        get => _offset;
        set {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentException("Panel border offset must be finite.", nameof(Offset));
            }
            _offset = value;
        }
    }

    /// <summary>Creates a copy of this panel border.</summary>
    public PdfPanelBorder Clone() => new PdfPanelBorder {
        Color = Color,
        Width = Width,
        Offset = Offset
    };
}
