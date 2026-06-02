using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable page border rendered as a page decoration.
/// </summary>
public sealed class PdfPageBorder {
    private double _width = 1D;
    private double _inset = 36D;
    private double _opacity = 1D;

    /// <summary>Border stroke color.</summary>
    public PdfColor Color { get; set; } = PdfColor.FromRgb(203, 213, 225);

    /// <summary>Border stroke width in points.</summary>
    public double Width {
        get => _width;
        set {
            Guard.Positive(value, nameof(Width));
            _width = value;
        }
    }

    /// <summary>Distance from the page edge to the border path, in points.</summary>
    public double Inset {
        get => _inset;
        set {
            Guard.NonNegative(value, nameof(Inset));
            _inset = value;
        }
    }

    /// <summary>Stroke opacity from 0 to 1. Defaults to 1.</summary>
    public double Opacity {
        get => _opacity;
        set {
            if (value < 0D || value > 1D || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentOutOfRangeException(nameof(Opacity), "PDF page border opacity must be a finite number between 0 and 1.");
            }

            _opacity = value;
        }
    }

    /// <summary>Border dash style.</summary>
    public OfficeStrokeDashStyle DashStyle { get; set; } = OfficeStrokeDashStyle.Solid;

    /// <summary>Creates a deep copy of this page border.</summary>
    public PdfPageBorder Clone() => new PdfPageBorder {
        Color = Color,
        Width = Width,
        Inset = Inset,
        Opacity = Opacity,
        DashStyle = DashStyle
    };
}
