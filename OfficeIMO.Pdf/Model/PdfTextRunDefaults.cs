namespace OfficeIMO.Pdf;

/// <summary>
/// Formatting that a <see cref="PdfParagraph"/> copies into each run it creates afterwards.
/// Changing a default never changes runs that already exist; format an existing run through its own setters.
/// </summary>
public sealed class PdfTextRunDefaults {
    private OfficeIMO.Drawing.OfficeTextDecorationStyle _underlineStyle;
    private OfficeIMO.Drawing.OfficeTextDecorationStyle _strikeStyle;
    private global::OfficeIMO.OfficeLength? _fontSize;
    private string? _fontFamily;
    private PdfStandardFont? _standardFont;
    private PdfTextBaseline _baseline;
    private OfficeIMO.Drawing.OfficeTextFeatureSettings _featureSettings = OfficeIMO.Drawing.OfficeTextFeatureSettings.Default;
    private double _horizontalTextScaling = 100D;

    internal PdfTextRunDefaults(global::OfficeIMO.OfficeColor? color) { Color = color; }

    /// <summary>Whether new runs are bold.</summary>
    public bool Bold { get; set; }

    /// <summary>Whether new runs are italic.</summary>
    public bool Italic { get; set; }

    /// <summary>Underline pattern for new runs. <see cref="OfficeIMO.Drawing.OfficeTextDecorationStyle.None"/> leaves them without underline.</summary>
    public OfficeIMO.Drawing.OfficeTextDecorationStyle UnderlineStyle {
        get => _underlineStyle;
        set { PdfTextRun.ValidateUnderlineStyle(value, nameof(UnderlineStyle)); _underlineStyle = value; }
    }

    /// <summary>Strikethrough pattern for new runs. <see cref="OfficeIMO.Drawing.OfficeTextDecorationStyle.None"/> leaves them without strikethrough.</summary>
    public OfficeIMO.Drawing.OfficeTextDecorationStyle StrikeStyle {
        get => _strikeStyle;
        set { PdfTextRun.ValidateStrikeStyle(value, nameof(StrikeStyle)); _strikeStyle = value; }
    }

    /// <summary>Foreground colour for new runs. Starts with the colour passed when the paragraph was added.</summary>
    public global::OfficeIMO.OfficeColor? Color { get; set; }

    /// <summary>Background colour for new runs, useful for highlights.</summary>
    public global::OfficeIMO.OfficeColor? BackgroundColor { get; set; }

    /// <summary>Font size for new runs. Null uses the paragraph font size.</summary>
    public global::OfficeIMO.OfficeLength? FontSize {
        get => _fontSize;
        set { if (value.HasValue) Guard.Positive(value.Value.ToPoints(), nameof(FontSize)); _fontSize = value; }
    }

    /// <summary>Registered embedded font family for new runs. Null uses the paragraph or document font.</summary>
    public string? FontFamily {
        get => _fontFamily;
        set { if (value != null) Guard.NotNullOrWhiteSpace(value, nameof(FontFamily)); _fontFamily = value?.Trim(); }
    }

    /// <summary>Standard PDF font for new runs. Null uses the paragraph or document font.</summary>
    public PdfStandardFont? StandardFont {
        get => _standardFont;
        set {
            if (value.HasValue) Guard.StandardFont(value.Value, nameof(StandardFont), "Text run font must be one of the supported standard PDF fonts.");
            _standardFont = value;
        }
    }

    /// <summary>Baseline placement for new runs.</summary>
    public PdfTextBaseline Baseline {
        get => _baseline;
        set { Guard.TextBaseline(value, nameof(Baseline)); _baseline = value; }
    }

    /// <summary>OpenType feature selections for new text runs, including kerning.</summary>
    public OfficeIMO.Drawing.OfficeTextFeatureSettings FeatureSettings {
        get => _featureSettings;
        set { Guard.NotNull(value, nameof(FeatureSettings)); _featureSettings = value; }
    }

    /// <summary>Horizontal glyph width percentage for new text runs. 100 is normal width.</summary>
    public double HorizontalTextScaling {
        get => _horizontalTextScaling;
        set { Guard.Positive(value, nameof(HorizontalTextScaling)); _horizontalTextScaling = value; }
    }

    /// <summary>Additional advance after each glyph of new text runs. Zero is normal spacing.</summary>
    public global::OfficeIMO.OfficeLength CharacterSpacing { get; set; }
}
