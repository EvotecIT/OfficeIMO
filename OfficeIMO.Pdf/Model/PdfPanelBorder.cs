namespace OfficeIMO.Pdf;

/// <summary>
/// Describes one side of a panel border.
/// </summary>
public sealed class PdfPanelBorder {
    private global::OfficeIMO.OfficeLength _width = global::OfficeIMO.OfficeLength.Points(0.5);
    private global::OfficeIMO.OfficeLength _offset;

    /// <summary>Border color. Set to null for no border on this side.</summary>
    public global::OfficeIMO.OfficeColor? Color { get; set; }

    /// <summary>Border stroke width in points.</summary>
    public global::OfficeIMO.OfficeLength Width {
        get => _width;
        set {
            if (value.ToPoints() < 0 || double.IsNaN(value.ToPoints()) || double.IsInfinity(value.ToPoints())) {
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
    public global::OfficeIMO.OfficeLength Offset {
        get => _offset;
        set {
            if (double.IsNaN(value.ToPoints()) || double.IsInfinity(value.ToPoints())) {
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
