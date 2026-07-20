namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable list style for bullet and numbered list typography and rhythm.
/// </summary>
public sealed class PdfListStyle {
    private double? _fontSize;
    private double? _lineHeight;
    private double _leftIndent;
    private double? _markerGap;
    private double? _markerWidth;
    private double? _markerFontSize;
    private string? _markerFontFamily;
    private PdfAlign? _markerAlign;
    private double _spacingBefore;
    private double? _spacingAfter;
    private double? _itemSpacing;

    /// <summary>List font size in points. When null the writer uses the current default font size.</summary>
    public double? FontSize {
        get => _fontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(FontSize), "List font size must be a positive finite value.");
            _fontSize = value;
        }
    }

    /// <summary>Line advance multiplier relative to the list font size. When null the writer uses the built-in list line height.</summary>
    public double? LineHeight {
        get => _lineHeight;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(LineHeight), "List line height must be a positive finite value.");
            _lineHeight = value;
        }
    }

    /// <summary>Horizontal inset from the list frame to the marker, in points.</summary>
    public double LeftIndent {
        get => _leftIndent;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(LeftIndent), "List left indent must be a non-negative finite value.");
            _leftIndent = value;
        }
    }

    /// <summary>Horizontal gap between the marker and text, in points. When null the writer uses one space advance.</summary>
    public double? MarkerGap {
        get => _markerGap;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(MarkerGap), "List marker gap must be a non-negative finite value.");
            _markerGap = value;
        }
    }

    /// <summary>Optional marker column width, in points. When null the writer estimates the width from the marker text.</summary>
    public double? MarkerWidth {
        get => _markerWidth;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(MarkerWidth), "List marker width must be a non-negative finite value.");
            _markerWidth = value;
        }
    }

    /// <summary>Vertical space before the list, in points.</summary>
    public double SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value, nameof(SpacingBefore), "List spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }

    /// <summary>Vertical space after the list, in points. When null the writer uses the item spacing.</summary>
    public double? SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(SpacingAfter), "List spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }

    /// <summary>Vertical space between list items, in points. When null the writer uses a small fraction of the line height.</summary>
    public double? ItemSpacing {
        get => _itemSpacing;
        set {
            ValidateOptionalNonNegativeFiniteValue(value, nameof(ItemSpacing), "List item spacing must be a non-negative finite value.");
            _itemSpacing = value;
        }
    }

    /// <summary>List marker and text color. A list block color overrides this value.</summary>
    public PdfColor? Color { get; set; }
    /// <summary>Optional list marker color. When null the marker uses the list block color or list text color.</summary>
    public PdfColor? MarkerColor { get; set; }
    /// <summary>Optional horizontal alignment of the marker inside the marker column. When null, bullets use left alignment and numbered lists use right alignment.</summary>
    public PdfAlign? MarkerAlign {
        get => _markerAlign;
        set {
            if (value.HasValue) {
                Guard.LeftCenterRightAlign(value.Value, nameof(MarkerAlign), "List marker");
            }

            _markerAlign = value;
        }
    }
    /// <summary>Optional standard font family used for list markers. When null the writer uses the current default font.</summary>
    public PdfStandardFont? MarkerFont { get; set; }
    /// <summary>Optional registered named font family used for list markers. When null, <see cref="MarkerFont"/> is used.</summary>
    public string? MarkerFontFamily {
        get => _markerFontFamily;
        set {
            if (value != null) {
                Guard.NotNullOrWhiteSpace(value, nameof(MarkerFontFamily));
            }

            _markerFontFamily = value?.Trim();
        }
    }
    /// <summary>Optional marker font size in points. When null the marker uses the list font size.</summary>
    public double? MarkerFontSize {
        get => _markerFontSize;
        set {
            ValidateOptionalPositiveFiniteValue(value, nameof(MarkerFontSize), "List marker font size must be a positive finite value.");
            _markerFontSize = value;
        }
    }
    /// <summary>When true, list markers render with the bold variant of the current list font.</summary>
    public bool MarkerBold { get; set; }
    /// <summary>When true, list markers render with the italic variant of the current list font.</summary>
    public bool MarkerItalic { get; set; }
    /// <summary>When true, the list moves as a unit instead of splitting across pages when it fits in the page frame.</summary>
    public bool KeepTogether { get; set; }
    /// <summary>When true, the list moves to the next page when it would otherwise be separated from the following flow block.</summary>
    public bool KeepWithNext { get; set; }

    /// <summary>Creates a copy of this list style.</summary>
    public PdfListStyle Clone() {
        return new PdfListStyle {
            FontSize = FontSize,
            LineHeight = LineHeight,
            LeftIndent = LeftIndent,
            MarkerGap = MarkerGap,
            MarkerWidth = MarkerWidth,
            SpacingBefore = SpacingBefore,
            SpacingAfter = SpacingAfter,
            ItemSpacing = ItemSpacing,
            Color = Color,
            MarkerColor = MarkerColor,
            MarkerAlign = MarkerAlign,
            MarkerFont = MarkerFont,
            MarkerFontFamily = MarkerFontFamily,
            MarkerFontSize = MarkerFontSize,
            MarkerBold = MarkerBold,
            MarkerItalic = MarkerItalic,
            KeepTogether = KeepTogether,
            KeepWithNext = KeepWithNext
        };
    }

    internal double GetFontSize(double defaultFontSize) {
        return FontSize ?? defaultFontSize;
    }

    internal double GetLeading(double fontSize) {
        return fontSize * (LineHeight ?? 1.4D);
    }

    internal double GetMarkerGap(double defaultGap) {
        return MarkerGap ?? defaultGap;
    }

    internal double GetItemSpacing(double leading) {
        return ItemSpacing ?? leading * 0.15D;
    }

    internal double GetSpacingAfter(double itemSpacing) {
        return SpacingAfter ?? itemSpacing;
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
