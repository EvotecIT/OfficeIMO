using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable visual and text defaults for fixed-position canvas text boxes.
/// </summary>
public sealed class PdfCanvasTextBoxStyle {
    private PdfAlign _align = PdfAlign.Left;
    private global::OfficeIMO.OfficeLength _borderWidth = global::OfficeIMO.OfficeLength.Points(0.5D);
    private global::OfficeIMO.OfficeLength _paddingX = global::OfficeIMO.OfficeLength.Points(6D);
    private global::OfficeIMO.OfficeLength _paddingY = global::OfficeIMO.OfficeLength.Points(4D);
    private global::OfficeIMO.OfficeLength? _paddingLeft;
    private global::OfficeIMO.OfficeLength? _paddingRight;
    private global::OfficeIMO.OfficeLength? _paddingTop;
    private global::OfficeIMO.OfficeLength? _paddingBottom;
    private global::OfficeIMO.OfficeLength _cornerRadius;
    private global::OfficeIMO.OfficeLength? _fontSize;
    private global::OfficeIMO.OfficeLength? _lineHeight;
    private double? _backgroundOpacity;
    private PdfStandardFont? _font;
    private PdfVerticalAlign _verticalAlign = PdfVerticalAlign.Top;

    // Format adapters can retain their native line-box baseline on fixed canvas text.
    internal PdfLineSpacing? LineSpacing { get; set; }

    /// <summary>Background fill color. Set to null for a transparent text box.</summary>
    public global::OfficeIMO.OfficeColor? Background { get; set; } = global::OfficeIMO.OfficeColor.White;

    /// <summary>Optional background fill opacity from 0 (transparent) to 1 (opaque).</summary>
    public double? BackgroundOpacity {
        get => _backgroundOpacity;
        set {
            ValidateOptionalOpacity(value, nameof(BackgroundOpacity));
            _backgroundOpacity = value;
        }
    }

    /// <summary>Border color. Set to null for no border.</summary>
    public global::OfficeIMO.OfficeColor? BorderColor { get; set; } = global::OfficeIMO.OfficeColor.FromNormalizedRgb(0.5D, 0.5D, 0.5D);

    /// <summary>Border stroke width in points.</summary>
    public global::OfficeIMO.OfficeLength BorderWidth {
        get => _borderWidth;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(BorderWidth), "Canvas text box border width must be a non-negative finite value.");
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
    public global::OfficeIMO.OfficeLength CornerRadius {
        get => _cornerRadius;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(CornerRadius), "Canvas text box corner radius must be a non-negative finite value.");
            _cornerRadius = value;
        }
    }

    /// <summary>Horizontal padding inside the text box, in points.</summary>
    public global::OfficeIMO.OfficeLength PaddingX {
        get => _paddingX;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(PaddingX), "Canvas text box horizontal padding must be a non-negative finite value.");
            _paddingX = value;
        }
    }

    /// <summary>Vertical padding inside the text box, in points.</summary>
    public global::OfficeIMO.OfficeLength PaddingY {
        get => _paddingY;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(PaddingY), "Canvas text box vertical padding must be a non-negative finite value.");
            _paddingY = value;
        }
    }

    /// <summary>Optional left padding override, in points. When null <see cref="PaddingX"/> is used.</summary>
    public global::OfficeIMO.OfficeLength? PaddingLeft {
        get => _paddingLeft;
        set {
            ValidateOptionalNonNegativeFiniteValue(value?.ToPoints(), nameof(PaddingLeft), "Canvas text box left padding must be a non-negative finite value.");
            _paddingLeft = value;
        }
    }

    /// <summary>Optional right padding override, in points. When null <see cref="PaddingX"/> is used.</summary>
    public global::OfficeIMO.OfficeLength? PaddingRight {
        get => _paddingRight;
        set {
            ValidateOptionalNonNegativeFiniteValue(value?.ToPoints(), nameof(PaddingRight), "Canvas text box right padding must be a non-negative finite value.");
            _paddingRight = value;
        }
    }

    /// <summary>Optional top padding override, in points. When null <see cref="PaddingY"/> is used.</summary>
    public global::OfficeIMO.OfficeLength? PaddingTop {
        get => _paddingTop;
        set {
            ValidateOptionalNonNegativeFiniteValue(value?.ToPoints(), nameof(PaddingTop), "Canvas text box top padding must be a non-negative finite value.");
            _paddingTop = value;
        }
    }

    /// <summary>Optional bottom padding override, in points. When null <see cref="PaddingY"/> is used.</summary>
    public global::OfficeIMO.OfficeLength? PaddingBottom {
        get => _paddingBottom;
        set {
            ValidateOptionalNonNegativeFiniteValue(value?.ToPoints(), nameof(PaddingBottom), "Canvas text box bottom padding must be a non-negative finite value.");
            _paddingBottom = value;
        }
    }

    internal double EffectivePaddingLeft => (PaddingLeft?.ToPoints()) ?? PaddingX.ToPoints();

    internal double EffectivePaddingRight => (PaddingRight?.ToPoints()) ?? PaddingX.ToPoints();

    internal double EffectivePaddingTop => (PaddingTop?.ToPoints()) ?? PaddingY.ToPoints();

    internal double EffectivePaddingBottom => (PaddingBottom?.ToPoints()) ?? PaddingY.ToPoints();

    /// <summary>Default foreground color for text runs that do not specify a color.</summary>
    public global::OfficeIMO.OfficeColor? TextColor { get; set; }

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
    public global::OfficeIMO.OfficeLength? FontSize {
        get => _fontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value?.ToPoints(), nameof(FontSize), "Canvas text box font size must be a positive finite value.");
            _fontSize = value;
        }
    }

    /// <summary>Default line height for wrapped text. When null, the font size is multiplied by 1.2.</summary>
    public global::OfficeIMO.OfficeLength? LineHeight {
        get => _lineHeight;
        set {
            ValidateOptionalPositiveFiniteValue(value?.ToPoints(), nameof(LineHeight), "Canvas text box line height must be a positive finite value.");
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
            BackgroundOpacity = BackgroundOpacity,
            BorderColor = BorderColor,
            BorderWidth = BorderWidth,
            BorderDashStyle = BorderDashStyle,
            BorderLineCap = BorderLineCap,
            BorderLineJoin = BorderLineJoin,
            CornerRadius = CornerRadius,
            PaddingX = PaddingX,
            PaddingY = PaddingY,
            PaddingLeft = PaddingLeft,
            PaddingRight = PaddingRight,
            PaddingTop = PaddingTop,
            PaddingBottom = PaddingBottom,
            TextColor = TextColor,
            Align = Align,
            VerticalAlign = VerticalAlign,
            FontSize = FontSize,
            LineHeight = LineHeight,
            LineSpacing = LineSpacing,
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

    private static void ValidateOptionalNonNegativeFiniteValue(double? value, string paramName, string message) {
        if (value.HasValue && (value.Value < 0 || double.IsNaN(value.Value) || double.IsInfinity(value.Value))) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidateOptionalOpacity(double? value, string paramName) {
        if (value.HasValue && (value.Value < 0D || value.Value > 1D || double.IsNaN(value.Value) || double.IsInfinity(value.Value))) {
            throw new System.ArgumentOutOfRangeException(paramName, "Canvas text box background opacity must be a finite value between 0 and 1.");
        }
    }

    private static void ValidateVerticalAlign(PdfVerticalAlign value, string paramName) {
        if (value != PdfVerticalAlign.Top && value != PdfVerticalAlign.Middle && value != PdfVerticalAlign.Bottom) {
            throw new System.ArgumentException("Canvas text box vertical alignment must be Top, Middle, or Bottom.", paramName);
        }
    }
}
