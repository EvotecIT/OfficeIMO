using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable visual and text defaults for fixed-position canvas text boxes.
/// </summary>
public sealed class PdfCanvasTextBoxStyle {
    private PdfAlign _align = PdfAlign.Left;
    private double _borderWidth = 0.5D;
    private double _paddingX = 6D;
    private double _paddingY = 4D;
    private double _cornerRadius;
    private double? _fontSize;
    private double? _lineHeight;
    private PdfStandardFont? _font;
    private PdfVerticalAlign _verticalAlign = PdfVerticalAlign.Top;

    /// <summary>Background fill color. Set to null for a transparent text box.</summary>
    public PdfColor? Background { get; set; } = PdfColor.White;

    /// <summary>Border color. Set to null for no border.</summary>
    public PdfColor? BorderColor { get; set; } = PdfColor.Gray;

    /// <summary>Border stroke width in points.</summary>
    public double BorderWidth {
        get => _borderWidth;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(BorderWidth), "Canvas text box border width must be a non-negative finite value.");
            _borderWidth = value;
        }
    }

    /// <summary>Border dash pattern.</summary>
    public OfficeStrokeDashStyle BorderDashStyle { get; set; } = OfficeStrokeDashStyle.Solid;

    /// <summary>Optional border line cap.</summary>
    public OfficeStrokeLineCap? BorderLineCap { get; set; }

    /// <summary>Optional border line join.</summary>
    public OfficeStrokeLineJoin? BorderLineJoin { get; set; }

    /// <summary>Corner radius for rounded text boxes. Zero renders a regular rectangle.</summary>
    public double CornerRadius {
        get => _cornerRadius;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(CornerRadius), "Canvas text box corner radius must be a non-negative finite value.");
            _cornerRadius = value;
        }
    }

    /// <summary>Horizontal padding inside the text box, in points.</summary>
    public double PaddingX {
        get => _paddingX;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(PaddingX), "Canvas text box horizontal padding must be a non-negative finite value.");
            _paddingX = value;
        }
    }

    /// <summary>Vertical padding inside the text box, in points.</summary>
    public double PaddingY {
        get => _paddingY;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(PaddingY), "Canvas text box vertical padding must be a non-negative finite value.");
            _paddingY = value;
        }
    }

    /// <summary>Default foreground color for text runs that do not specify a color.</summary>
    public PdfColor? TextColor { get; set; }

    /// <summary>Text alignment inside the text box.</summary>
    public PdfAlign Align {
        get => _align;
        set {
            Guard.ParagraphAlign(value, nameof(Align), "Canvas text box");
            _align = value;
        }
    }

    /// <summary>Vertical alignment of wrapped text inside the text box.</summary>
    public PdfVerticalAlign VerticalAlign {
        get => _verticalAlign;
        set {
            ValidateVerticalAlign(value, nameof(VerticalAlign));
            _verticalAlign = value;
        }
    }

    /// <summary>Default font size for text runs that do not specify a font size.</summary>
    public double? FontSize {
        get => _fontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(FontSize), "Canvas text box font size must be a positive finite value.");
            _fontSize = value;
        }
    }

    /// <summary>Default line height for wrapped text. When null, the font size is multiplied by 1.2.</summary>
    public double? LineHeight {
        get => _lineHeight;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(LineHeight), "Canvas text box line height must be a positive finite value.");
            _lineHeight = value;
        }
    }

    /// <summary>Default standard PDF font for text runs that do not specify a font.</summary>
    public PdfStandardFont? Font {
        get => _font;
        set {
            if (value.HasValue) {
                Guard.StandardFont(value.Value, nameof(Font), "Canvas text box font must be one of the supported standard PDF fonts.");
            }

            _font = value;
        }
    }

    /// <summary>Creates a copy of this canvas text box style.</summary>
    public PdfCanvasTextBoxStyle Clone() {
        return new PdfCanvasTextBoxStyle {
            Background = Background,
            BorderColor = BorderColor,
            BorderWidth = BorderWidth,
            BorderDashStyle = BorderDashStyle,
            BorderLineCap = BorderLineCap,
            BorderLineJoin = BorderLineJoin,
            CornerRadius = CornerRadius,
            PaddingX = PaddingX,
            PaddingY = PaddingY,
            TextColor = TextColor,
            Align = Align,
            VerticalAlign = VerticalAlign,
            FontSize = FontSize,
            LineHeight = LineHeight,
            Font = Font
        };
    }

    private static void ValidateNonNegativeFiniteValue(double value, string paramName, string message) {
        if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidateOptionalPositiveFiniteValue(double? value, string paramName, string message) {
        if (value.HasValue && (value.Value <= 0 || double.IsNaN(value.Value) || double.IsInfinity(value.Value))) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidateVerticalAlign(PdfVerticalAlign value, string paramName) {
        if (value != PdfVerticalAlign.Top && value != PdfVerticalAlign.Middle && value != PdfVerticalAlign.Bottom) {
            throw new System.ArgumentException("Canvas text box vertical alignment must be Top, Middle, or Bottom.", paramName);
        }
    }
}
