namespace OfficeIMO.Pdf;

public sealed partial class PdfOptions {
    /// <summary>Default standard font used for paragraphs.</summary>
    public PdfStandardFont DefaultFont {
        get => _defaultFont;
        set {
            Guard.StandardFont(value, nameof(DefaultFont), "PDF default font must be one of the supported standard PDF fonts.");
            _defaultFont = value;
            _hasExplicitDefaultFont = true;
        }
    }
    /// <summary>Gets whether the default paragraph font slot was explicitly supplied by the caller or a theme.</summary>
    public bool HasExplicitDefaultFont => _hasExplicitDefaultFont;
    /// <summary>Default paragraph font size in points. Default 11.</summary>
    public double DefaultFontSize { get; set; } = 11;
    /// <summary>Default text color for blocks when none is specified.</summary>
    public PdfColor? DefaultTextColor { get; set; }
    /// <summary>Default paragraph style applied when a paragraph does not specify its own style.</summary>
    public PdfParagraphStyle? DefaultParagraphStyle {
        get => _defaultParagraphStyle?.Clone();
        set => _defaultParagraphStyle = value?.Clone();
    }
    /// <summary>Default table style applied when none is provided.</summary>
    public PdfTableStyle? DefaultTableStyle {
        get => _defaultTableStyle?.Clone();
        set {
            _defaultTableStyle = value?.Clone();
            _hasExplicitDefaultTableStyle = value != null;
        }
    }
    /// <summary>Gets whether <see cref="DefaultTableStyle"/> was explicitly supplied by the caller or a theme.</summary>
    public bool HasExplicitDefaultTableStyle => _hasExplicitDefaultTableStyle;
    /// <summary>Default heading styles applied when H1/H2/H3 blocks do not specify their own style.</summary>
    public PdfHeadingStyles? DefaultHeadingStyles {
        get => _defaultHeadingStyles?.Clone();
        set => _defaultHeadingStyles = value?.Clone();
    }
    /// <summary>Default list style applied when bullet and numbered lists do not specify their own style.</summary>
    public PdfListStyle? DefaultListStyle {
        get => _defaultListStyle?.Clone();
        set => _defaultListStyle = value?.Clone();
    }
    /// <summary>Default panel style applied when panel paragraphs do not specify their own style.</summary>
    public PanelStyle? DefaultPanelStyle {
        get => _defaultPanelStyle?.Clone();
        set => _defaultPanelStyle = value?.Clone();
    }
    /// <summary>Default horizontal rule style applied when horizontal rules do not specify their own style.</summary>
    public PdfHorizontalRuleStyle? DefaultHorizontalRuleStyle {
        get => _defaultHorizontalRuleStyle?.Clone();
        set => _defaultHorizontalRuleStyle = value?.Clone();
    }
    /// <summary>Default image placement style applied when images do not specify their own style.</summary>
    public PdfImageStyle? DefaultImageStyle {
        get => _defaultImageStyle?.Clone();
        set => _defaultImageStyle = value?.Clone();
    }
    /// <summary>Default placement style for OfficeIMO.Drawing-backed flow objects.</summary>
    public PdfDrawingStyle? DefaultDrawingStyle {
        get => _defaultDrawingStyle?.Clone();
        set => _defaultDrawingStyle = value?.Clone();
    }
    /// <summary>Default row/column layout style applied when rows do not specify their own style.</summary>
    public PdfRowStyle? DefaultRowStyle {
        get => _defaultRowStyle?.Clone();
        set => _defaultRowStyle = value?.Clone();
    }
    /// <summary>Optional debug overlays (margins, baselines, boxes).</summary>
    public PdfDebugOptions? Debug { get; set; }

    /// <summary>When true, H1/H2/H3 blocks are written as PDF outline/bookmark entries.</summary>
    public bool CreateOutlineFromHeadings { get; set; }

    /// <summary>
    /// Highest outline level expanded when generated heading outlines are opened in a PDF reader. Defaults to all levels.
    /// Set to 0 to show only top-level entries with children collapsed.
    /// </summary>
    public int OutlineExpansionLevel {
        get => _outlineExpansionLevel;
        set {
            if (value < 0) {
                throw new ArgumentOutOfRangeException(nameof(value), value, "PDF outline expansion level must be non-negative.");
            }

            _outlineExpansionLevel = value;
        }
    }

    /// <summary>Applies reusable default styles to this options object.</summary>
    public PdfOptions ApplyTheme(PdfTheme theme) {
        Guard.NotNull(theme, nameof(theme));
        theme.Clone().ApplyTo(this);
        return this;
    }
    internal PdfParagraphStyle? DefaultParagraphStyleSnapshot => _defaultParagraphStyle;
    internal PdfTableStyle? DefaultTableStyleSnapshot => _defaultTableStyle;
    internal bool HasExplicitDefaultTableStyleSnapshot => _hasExplicitDefaultTableStyle;
    internal PdfHeadingStyles? DefaultHeadingStylesSnapshot => _defaultHeadingStyles;
    internal PdfListStyle? DefaultListStyleSnapshot => _defaultListStyle;
    internal PanelStyle? DefaultPanelStyleSnapshot => _defaultPanelStyle;
    internal PdfHorizontalRuleStyle? DefaultHorizontalRuleStyleSnapshot => _defaultHorizontalRuleStyle;
    internal PdfImageStyle? DefaultImageStyleSnapshot => _defaultImageStyle;
    internal PdfDrawingStyle? DefaultDrawingStyleSnapshot => _defaultDrawingStyle;
    internal PdfRowStyle? DefaultRowStyleSnapshot => _defaultRowStyle;
    internal int OutlineExpansionLevelSnapshot => _outlineExpansionLevel;

    /// <summary>Sets the default style for a built-in heading level.</summary>
    public PdfOptions SetDefaultHeadingStyle(int level, PdfHeadingStyle style) {
        Guard.NotNull(style, nameof(style));
        (_defaultHeadingStyles ??= new PdfHeadingStyles()).Set(level, style);
        return this;
    }

}
