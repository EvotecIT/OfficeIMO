namespace OfficeIMO.Pdf;

/// <summary>
/// Visual style of a panel box used by panel paragraphs.
/// </summary>
public class PanelStyle {
    private PdfAlign _align = PdfAlign.Left;
    private double _borderWidth = 0.5;
    private double _paddingY = 6;
    private double _paddingX = 6;
    private double? _maxWidth;
    private double _spacingBefore;
    private double _spacingAfter = 6;

    /// <summary>Background fill color. Set to null for no fill.</summary>
    public PdfColor? Background { get; set; }
    /// <summary>Border color. Set to null for no border.</summary>
    public PdfColor? BorderColor { get; set; }
    /// <summary>Border stroke width in points.</summary>
    public double BorderWidth {
        get => _borderWidth;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(BorderWidth), "Panel border width must be a non-negative finite value.");
            _borderWidth = value;
        }
    }
    /// <summary>Vertical padding inside the panel (points).</summary>
    public double PaddingY {
        get => _paddingY;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(PaddingY), "Panel vertical padding must be a non-negative finite value.");
            _paddingY = value;
        }
    }
    /// <summary>Horizontal padding inside the panel (points).</summary>
    public double PaddingX {
        get => _paddingX;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(PaddingX), "Panel horizontal padding must be a non-negative finite value.");
            _paddingX = value;
        }
    }
    /// <summary>Optional maximum width for the panel box (points). When set, the box can be centered or right-aligned.</summary>
    public double? MaxWidth {
        get => _maxWidth;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(MaxWidth), "Panel maximum width must be a positive finite value.");
            _maxWidth = value;
        }
    }
    /// <summary>Horizontal alignment of the panel box within the content area.</summary>
    public PdfAlign Align {
        get => _align;
        set {
            Guard.LeftCenterRightAlign(value, nameof(Align), "Panel box");
            _align = value;
        }
    }
    /// <summary>Vertical space before the panel in the surrounding document flow, in points.</summary>
    public double SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingBefore), "Panel spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }
    /// <summary>Vertical space after the panel in the surrounding document flow, in points.</summary>
    public double SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingAfter), "Panel spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }
    /// <summary>When true, the entire panel is kept on one page; otherwise it can split across pages.</summary>
    public bool KeepTogether { get; set; }
    /// <summary>When true, the panel moves to the next page when it would otherwise be separated from the following flow block.</summary>
    public bool KeepWithNext { get; set; }

    /// <summary>Creates a copy of this panel style.</summary>
    public PanelStyle Clone() {
        return new PanelStyle {
            Background = Background,
            BorderColor = BorderColor,
            BorderWidth = BorderWidth,
            PaddingY = PaddingY,
            PaddingX = PaddingX,
            MaxWidth = MaxWidth,
            Align = Align,
            SpacingBefore = SpacingBefore,
            SpacingAfter = SpacingAfter,
            KeepTogether = KeepTogether,
            KeepWithNext = KeepWithNext
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
}
