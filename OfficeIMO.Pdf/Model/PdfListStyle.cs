namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable list style for bullet and numbered list typography and rhythm.
/// </summary>
public sealed class PdfListStyle {
    private global::OfficeIMO.OfficeLength? _fontSize;
    private double? _lineHeight;

    /// <summary>Explicit line spacing. When set, this overrides <see cref="LineHeight"/>.</summary>
    public PdfLineSpacing? LineSpacing { get; set; }
    private global::OfficeIMO.OfficeLength _leftIndent;
    private global::OfficeIMO.OfficeLength? _markerGap;
    private global::OfficeIMO.OfficeLength? _markerWidth;
    private global::OfficeIMO.OfficeLength? _markerFontSize;
    private string? _markerFontFamily;
    private PdfAlign? _markerAlign;
    private global::OfficeIMO.OfficeLength _spacingBefore;
    private global::OfficeIMO.OfficeLength? _spacingAfter;
    private global::OfficeIMO.OfficeLength? _itemSpacing;

    // Word numbering aligns marker ink at the hanging-indent anchor. Ordinary
    // PDF lists continue to align within their reserved marker column.
    internal bool MarkerAlignsAtIndent { get; set; }

    // Format converters can supply marker shaping separately from body text.
    internal OfficeIMO.Drawing.OfficeTextFeatureSettings? MarkerFeatureSettings { get; set; }

    /// <summary>List font size in points. When null the writer uses the current default font size.</summary>
    public global::OfficeIMO.OfficeLength? FontSize {
        get => _fontSize;
        set {
            ValidateOptionalPositiveFiniteValue((value?.ToPoints()), nameof(FontSize), "List font size must be a positive finite value.");
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
    public global::OfficeIMO.OfficeLength LeftIndent {
        get => _leftIndent;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(LeftIndent), "List left indent must be a non-negative finite value.");
            _leftIndent = value;
        }
    }

    /// <summary>Horizontal gap between the marker and text, in points. When null the writer uses one space advance.</summary>
    public global::OfficeIMO.OfficeLength? MarkerGap {
        get => _markerGap;
        set {
            ValidateOptionalNonNegativeFiniteValue((value?.ToPoints()), nameof(MarkerGap), "List marker gap must be a non-negative finite value.");
            _markerGap = value;
        }
    }

    /// <summary>Optional marker column width, in points. When null the writer estimates the width from the marker text.</summary>
    public global::OfficeIMO.OfficeLength? MarkerWidth {
        get => _markerWidth;
        set {
            ValidateOptionalNonNegativeFiniteValue((value?.ToPoints()), nameof(MarkerWidth), "List marker width must be a non-negative finite value.");
            _markerWidth = value;
        }
    }

    /// <summary>Vertical space before the list, in points.</summary>
    public global::OfficeIMO.OfficeLength SpacingBefore {
        get => _spacingBefore;
        set {
            ValidateNonNegativeFiniteValue(value.ToPoints(), nameof(SpacingBefore), "List spacing before must be a non-negative finite value.");
            _spacingBefore = value;
        }
    }

    /// <summary>Vertical space after the list, in points. When null the writer uses the item spacing.</summary>
    public global::OfficeIMO.OfficeLength? SpacingAfter {
        get => _spacingAfter;
        set {
            ValidateOptionalNonNegativeFiniteValue((value?.ToPoints()), nameof(SpacingAfter), "List spacing after must be a non-negative finite value.");
            _spacingAfter = value;
        }
    }

    /// <summary>Vertical space between list items, in points. When null the writer uses a small fraction of the line height.</summary>
    public global::OfficeIMO.OfficeLength? ItemSpacing {
        get => _itemSpacing;
        set {
            ValidateOptionalNonNegativeFiniteValue((value?.ToPoints()), nameof(ItemSpacing), "List item spacing must be a non-negative finite value.");
            _itemSpacing = value;
        }
    }

    /// <summary>List marker and text color. A list block color overrides this value.</summary>
    public global::OfficeIMO.OfficeColor? Color { get; set; }
    /// <summary>Optional list marker color. When null the marker uses the list block color or list text color.</summary>
    public global::OfficeIMO.OfficeColor? MarkerColor { get; set; }
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
    public global::OfficeIMO.OfficeLength? MarkerFontSize {
        get => _markerFontSize;
        set {
            ValidateOptionalPositiveFiniteValue((value?.ToPoints()), nameof(MarkerFontSize), "List marker font size must be a positive finite value.");
            _markerFontSize = value;
        }
    }
    /// <summary>When true, list markers render with the bold variant of the current list font.</summary>
    public bool MarkerBold { get; set; }
    /// <summary>When true, list markers render with the italic variant of the current list font.</summary>
    public bool MarkerItalic { get; set; }
    /// <summary>When true, the list moves as a unit instead of splitting across pages when it fits in the page frame.</summary>
    public bool KeepTogether { get; set; }
    /// <summary>Imported document policy for a kept list paragraph taller than a complete continuation frame.</summary>
    internal PdfKeepTogetherOverflowBehavior KeepTogetherOverflow { get; set; }
    /// <summary>When true, the list moves to the next page when it would otherwise be separated from the following flow block.</summary>
    public bool KeepWithNext { get; set; }

    /// <summary>Creates a copy of this list style.</summary>
    public PdfListStyle Clone() {
        return new PdfListStyle {
            FontSize = FontSize,
            LineSpacing = LineSpacing,
            LineHeight = LineHeight,
            LeftIndent = LeftIndent,
            MarkerGap = MarkerGap,
            MarkerWidth = MarkerWidth,
            MarkerAlignsAtIndent = MarkerAlignsAtIndent,
            MarkerFeatureSettings = MarkerFeatureSettings,
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
            KeepTogetherOverflow = KeepTogetherOverflow,
            KeepWithNext = KeepWithNext
        };
    }

    internal double GetFontSize(double defaultFontSize) {
        return (FontSize?.ToPoints()) ?? defaultFontSize;
    }

    internal double GetLeading(double fontSize) {
        return LineSpacing?.GetAdvance(fontSize) ?? fontSize * (LineHeight ?? 1.4D);
    }

    internal double GetMarkerGap(double defaultGap) {
        return (MarkerGap?.ToPoints()) ?? defaultGap;
    }

    internal double GetItemSpacing(double leading) {
        return (ItemSpacing?.ToPoints()) ?? leading * 0.15D;
    }

    internal double GetSpacingAfter(double itemSpacing) {
        return (SpacingAfter?.ToPoints()) ?? itemSpacing;
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
