namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable heading style for H1/H2/H3 typography and vertical rhythm.
/// </summary>
public sealed class PdfHeadingStyle {
    private double? _fontSize;
    private double? _lineHeight;
    private double _spacingBefore;
    private double? _spacingAfter;
    private string? _fontFamily;

    /// <summary>Heading font size in points. When null the writer uses the built-in size for the heading level.</summary>
    public double? FontSize {
        get => _fontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(FontSize), "Heading font size must be a positive finite value.");
            _fontSize = value;
        }
    }

    /// <summary>Line advance multiplier relative to the heading font size. When null the writer uses the built-in heading line height.</summary>
    public double? LineHeight {
        get => _lineHeight;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(LineHeight), "Heading line height must be a positive finite value.");
            _lineHeight = value;
        }
    }

    /// <summary>Vertical space before the heading, in points.</summary>
    public double SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingBefore), "Heading spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }

    /// <summary>Vertical space after the heading, in points. When null the writer uses the built-in heading spacing.</summary>
    public double? SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(SpacingAfter), "Heading spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }

    /// <summary>Heading text color. A heading block color overrides this value.</summary>
    public PdfColor? Color { get; set; }

    /// <summary>Heading font slot. When null the writer uses the document default font family.</summary>
    public PdfStandardFont? Font { get; set; }

    /// <summary>Optional registered embedded family used by heading text. <see cref="Font"/> remains its fallback.</summary>
    public string? FontFamily {
        get => _fontFamily;
        set {
            if (value != null) {
                Guard.NotNullOrWhiteSpace(value, nameof(FontFamily));
            }

            _fontFamily = value?.Trim();
        }
    }

    /// <summary>When true, headings use the bold variant of the document font.</summary>
    public bool Bold { get; set; } = true;

    /// <summary>When true, <see cref="SpacingBefore"/> is honored even when the heading starts a page or column.</summary>
    public bool ApplySpacingBeforeAtTop { get; set; }

    /// <summary>When true, the heading moves to the next page when it would otherwise be separated from the following paragraph.</summary>
    public bool KeepWithNext { get; set; } = true;

    /// <summary>Creates a copy of this heading style.</summary>
    public PdfHeadingStyle Clone() {
        return new PdfHeadingStyle {
            FontSize = FontSize,
            LineHeight = LineHeight,
            SpacingBefore = SpacingBefore,
            SpacingAfter = SpacingAfter,
            Color = Color,
            Font = Font,
            FontFamily = FontFamily,
            Bold = Bold,
            ApplySpacingBeforeAtTop = ApplySpacingBeforeAtTop,
            KeepWithNext = KeepWithNext
        };
    }

    internal double GetFontSize(int level) {
        return FontSize ?? GetDefaultFontSize(level);
    }

    internal double GetLeading(double fontSize) {
        return fontSize * (LineHeight ?? 1.25D);
    }

    internal double GetSpacingAfter(double leading) {
        return SpacingAfter ?? leading * 0.25D;
    }

    internal static double GetDefaultFontSize(int level) {
        return level switch { 1 => 24D, 2 => 18D, 3 => 14D, _ => 12D };
    }

    private static void ValidateNonNegativeFiniteValue(double value, string paramName, string message) {
        if (value < 0 || double.IsNaN(value) || double.IsInfinity(value)) {
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
