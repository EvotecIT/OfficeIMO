namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable text style for document or page default typography.
/// </summary>
public sealed class PdfTextStyle {
    private PdfStandardFont _font = PdfStandardFont.Helvetica;
    private double _fontSize = 11;

    /// <summary>Default standard font family.</summary>
    public PdfStandardFont Font {
        get => _font;
        set {
            Guard.StandardFont(value, nameof(Font), "PDF text style font must be one of the supported standard PDF fonts.");
            _font = value;
        }
    }

    /// <summary>Default font size in points.</summary>
    public double FontSize {
        get => _fontSize;
        set {
            Guard.Positive(value, nameof(FontSize));
            _fontSize = value;
        }
    }

    /// <summary>Default text color. When null, the writer default is used.</summary>
    public PdfColor? Color { get; set; }

    /// <summary>Creates a deep copy of this text style.</summary>
    public PdfTextStyle Clone() {
        return new PdfTextStyle {
            Font = Font,
            FontSize = FontSize,
            Color = Color
        };
    }

    internal void ApplyTo(PdfOptions options) {
        Guard.NotNull(options, nameof(options));
        options.DefaultFont = Font;
        options.DefaultFontSize = FontSize;
        options.DefaultTextColor = Color;
    }
}
