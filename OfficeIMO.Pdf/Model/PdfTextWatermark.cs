namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable text watermark rendered behind page content.
/// </summary>
public sealed class PdfTextWatermark {
    private string _text = string.Empty;
    private double _fontSize = 54D;
    private PdfStandardFont _font = PdfStandardFont.Helvetica;
    private double _opacity = 0.12D;
    private double _rotationAngle = -35D;

    /// <summary>Creates a text watermark.</summary>
    public PdfTextWatermark(string text) {
        Text = text;
    }

    /// <summary>Watermark text.</summary>
    public string Text {
        get => _text;
        set {
            Guard.NotNullOrWhiteSpace(value, nameof(Text));
            _text = value;
        }
    }

    /// <summary>Standard PDF font family used for the watermark.</summary>
    public PdfStandardFont Font {
        get => _font;
        set {
            Guard.StandardFont(value, nameof(Font), "PDF watermark font must be one of the supported standard PDF fonts.");
            _font = value;
        }
    }

    /// <summary>Watermark font size in points.</summary>
    public double FontSize {
        get => _fontSize;
        set {
            Guard.Positive(value, nameof(FontSize));
            _fontSize = value;
        }
    }

    /// <summary>Watermark fill color. Defaults to a neutral gray.</summary>
    public PdfColor Color { get; set; } = PdfColor.FromRgb(148, 163, 184);

    /// <summary>Fill opacity from 0 to 1. Defaults to 0.12.</summary>
    public double Opacity {
        get => _opacity;
        set {
            if (value < 0D || value > 1D || double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentOutOfRangeException(nameof(Opacity), "PDF watermark opacity must be a finite number between 0 and 1.");
            }

            _opacity = value;
        }
    }

    /// <summary>Rotation angle in degrees. Defaults to -35.</summary>
    public double RotationAngle {
        get => _rotationAngle;
        set {
            if (double.IsNaN(value) || double.IsInfinity(value)) {
                throw new System.ArgumentOutOfRangeException(nameof(RotationAngle), "PDF watermark rotation angle must be finite.");
            }

            _rotationAngle = value;
        }
    }

    /// <summary>Use the bold variant of <see cref="Font"/> when available.</summary>
    public bool Bold { get; set; } = true;

    /// <summary>Use the italic variant of <see cref="Font"/> when available.</summary>
    public bool Italic { get; set; }

    /// <summary>Creates a deep copy of this watermark.</summary>
    public PdfTextWatermark Clone() => new PdfTextWatermark(Text) {
        Font = Font,
        FontSize = FontSize,
        Color = Color,
        Opacity = Opacity,
        RotationAngle = RotationAngle,
        Bold = Bold,
        Italic = Italic
    };
}
