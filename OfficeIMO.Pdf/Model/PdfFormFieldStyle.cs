namespace OfficeIMO.Pdf;

/// <summary>
/// Visual style for generated simple AcroForm fields.
/// </summary>
public class PdfFormFieldStyle {
    private double _borderWidth = 1D;

    /// <summary>Background fill color. Set to null for transparent field appearance streams.</summary>
    public PdfColor? BackgroundColor { get; set; } = PdfColor.White;

    /// <summary>Border stroke color. Set to null for no border stroke.</summary>
    public PdfColor? BorderColor { get; set; } = new PdfColor(0.75, 0.75, 0.75);

    /// <summary>Text color for generated text and choice field appearance streams.</summary>
    public PdfColor TextColor { get; set; } = PdfColor.Black;

    /// <summary>Check mark or radio dot color for generated button field appearance streams.</summary>
    public PdfColor MarkColor { get; set; } = PdfColor.Black;

    /// <summary>Border stroke width in points. Set to 0 to suppress border drawing.</summary>
    public double BorderWidth {
        get => _borderWidth;
        set {
            if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new ArgumentOutOfRangeException(nameof(value), value, "PDF form field border width must be a non-negative finite number.");
            }

            _borderWidth = value;
        }
    }

    /// <summary>Creates a copy of this form field style.</summary>
    public PdfFormFieldStyle Clone() {
        return new PdfFormFieldStyle {
            BackgroundColor = BackgroundColor,
            BorderColor = BorderColor,
            BorderWidth = BorderWidth,
            TextColor = TextColor,
            MarkColor = MarkColor
        };
    }
}
