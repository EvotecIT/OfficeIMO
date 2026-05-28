namespace OfficeIMO.Pdf;

/// <summary>
/// Describes layout options for rich paragraph rendering.
/// </summary>
public class PdfParagraphStyle {
    private double? _lineHeight;
    private double _leftIndent;
    private double _rightIndent;
    private double _firstLineIndent;
    private double _spacingBefore;
    private double? _spacingAfter;
    private double? _defaultTabStopWidth;

    /// <summary>Line advance multiplier relative to the paragraph font size. When null the writer uses the default line height.</summary>
    public double? LineHeight {
        get => _lineHeight;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(LineHeight), "Paragraph line height must be a positive finite value.");
            _lineHeight = value;
        }
    }
    /// <summary>Horizontal inset from the left edge of the paragraph frame, in points.</summary>
    public double LeftIndent {
        get => _leftIndent;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(LeftIndent), "Paragraph left indent must be a non-negative finite value.");
            _leftIndent = value;
        }
    }
    /// <summary>Horizontal inset from the right edge of the paragraph frame, in points.</summary>
    public double RightIndent {
        get => _rightIndent;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(RightIndent), "Paragraph right indent must be a non-negative finite value.");
            _rightIndent = value;
        }
    }
    /// <summary>Additional indent for the first line, in points. Negative values create a hanging indent relative to <see cref="LeftIndent"/>.</summary>
    public double FirstLineIndent {
        get => _firstLineIndent;
        set {
            ValidateFiniteValue(value, nameof(FirstLineIndent), "Paragraph first line indent must be a finite value.");
            _firstLineIndent = value;
        }
    }
    /// <summary>Vertical space before the paragraph, in points.</summary>
    public double SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingBefore), "Paragraph spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }
    /// <summary>Vertical space after the paragraph, in points. When null the writer uses the default paragraph spacing.</summary>
    public double? SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(SpacingAfter), "Paragraph spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }
    /// <summary>Default paragraph tab-stop width, in points. When null the writer uses the Word-compatible half-inch default.</summary>
    public double? DefaultTabStopWidth {
        get => _defaultTabStopWidth;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(DefaultTabStopWidth), "Paragraph default tab stop width must be a positive finite value.");
            _defaultTabStopWidth = value;
        }
    }
    /// <summary>When true, the paragraph starts on a new page instead of splitting across pages.</summary>
    public bool KeepTogether { get; set; }
    /// <summary>When true, the paragraph moves to the next page when it would otherwise be separated from the following paragraph.</summary>
    public bool KeepWithNext { get; set; }
    /// <summary>When true, paragraph page splits avoid single-line widows and orphans where the page frame allows it.</summary>
    public bool WidowControl { get; set; }

    /// <summary>Creates a copy of this paragraph style.</summary>
    public PdfParagraphStyle Clone() {
        return new PdfParagraphStyle {
            LineHeight = LineHeight,
            LeftIndent = LeftIndent,
            RightIndent = RightIndent,
            FirstLineIndent = FirstLineIndent,
            SpacingBefore = SpacingBefore,
            SpacingAfter = SpacingAfter,
            DefaultTabStopWidth = DefaultTabStopWidth,
            KeepTogether = KeepTogether,
            KeepWithNext = KeepWithNext,
            WidowControl = WidowControl
        };
    }

    private static void ValidateNonNegativeFiniteValue(double value, string paramName, string message) {
        if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidateFiniteValue(double value, string paramName, string message) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new System.ArgumentException(message, paramName);
        }
    }

    private static void ValidateOptionalNonNegativeFiniteValue(double? value, string paramName, string message) {
        if (value.HasValue) {
            ValidateNonNegativeFiniteValue(value.Value, paramName, message);
        }
    }

    private static void ValidateOptionalPositiveFiniteValue(double? value, string paramName, string message) {
        if (value.HasValue && (value.Value <= 0 || double.IsNaN(value.Value) || double.IsInfinity(value.Value))) {
            throw new System.ArgumentException(message, paramName);
        }
    }
}
