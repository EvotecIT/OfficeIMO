namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable bundle of default PDF styles for document or page composition.
/// </summary>
public sealed class PdfTheme {
    private PdfTextStyle? _textStyle;
    private PdfParagraphStyle? _paragraphStyle;
    private PdfTableStyle? _tableStyle;
    private PdfHeadingStyles? _headingStyles;
    private PdfListStyle? _listStyle;
    private PanelStyle? _panelStyle;
    private PdfHorizontalRuleStyle? _horizontalRuleStyle;
    private PdfImageStyle? _imageStyle;
    private PdfDrawingStyle? _drawingStyle;
    private PdfRowStyle? _rowStyle;

    /// <summary>Default text typography for content that does not override it.</summary>
    public PdfTextStyle? TextStyle {
        get => _textStyle?.Clone();
        set => _textStyle = value?.Clone();
    }

    /// <summary>Default paragraph layout and page-flow style.</summary>
    public PdfParagraphStyle? ParagraphStyle {
        get => _paragraphStyle?.Clone();
        set => _paragraphStyle = value?.Clone();
    }

    /// <summary>Default table appearance and rhythm style.</summary>
    public PdfTableStyle? TableStyle {
        get => _tableStyle?.Clone();
        set => _tableStyle = value?.Clone();
    }

    /// <summary>Default heading typography and rhythm styles.</summary>
    public PdfHeadingStyles? HeadingStyles {
        get => _headingStyles?.Clone();
        set => _headingStyles = value?.Clone();
    }

    /// <summary>Default bullet and numbered list typography and rhythm style.</summary>
    public PdfListStyle? ListStyle {
        get => _listStyle?.Clone();
        set => _listStyle = value?.Clone();
    }

    /// <summary>Default panel paragraph box style.</summary>
    public PanelStyle? PanelStyle {
        get => _panelStyle?.Clone();
        set => _panelStyle = value?.Clone();
    }

    /// <summary>Default horizontal rule appearance and rhythm style.</summary>
    public PdfHorizontalRuleStyle? HorizontalRuleStyle {
        get => _horizontalRuleStyle?.Clone();
        set => _horizontalRuleStyle = value?.Clone();
    }

    /// <summary>Default image placement and rhythm style.</summary>
    public PdfImageStyle? ImageStyle {
        get => _imageStyle?.Clone();
        set => _imageStyle = value?.Clone();
    }

    /// <summary>Default drawing object placement and rhythm style.</summary>
    public PdfDrawingStyle? DrawingStyle {
        get => _drawingStyle?.Clone();
        set => _drawingStyle = value?.Clone();
    }

    /// <summary>Default row/column layout rhythm style.</summary>
    public PdfRowStyle? RowStyle {
        get => _rowStyle?.Clone();
        set => _rowStyle = value?.Clone();
    }

    /// <summary>
    /// Creates a generic Word-like document theme with neutral typography, readable paragraph rhythm, heading hierarchy, lists, tables, and flow-object spacing.
    /// </summary>
    public static PdfTheme WordLike() {
        PdfColor bodyText = PdfColor.FromRgb(31, 41, 55);
        PdfColor headingText = PdfColor.FromRgb(17, 24, 39);
        PdfColor ruleColor = PdfColor.FromRgb(203, 213, 225);

        var tableStyle = TableStyles.ListTable1Light();
        tableStyle.TextColor = bodyText;
        tableStyle.HeaderTextColor = headingText;
        tableStyle.CellPaddingX = 5;
        tableStyle.CellPaddingY = 5;
        tableStyle.RowSeparatorColor = PdfColor.FromRgb(226, 232, 240);
        tableStyle.HeaderSeparatorColor = headingText;
        tableStyle.FooterSeparatorColor = headingText;
        tableStyle.FooterSeparatorWidth = 0.8;

        return new PdfTheme {
            TextStyle = new PdfTextStyle {
                Font = PdfStandardFont.Helvetica,
                FontSize = 11,
                Color = bodyText
            },
            ParagraphStyle = new PdfParagraphStyle {
                LineHeight = 1.15,
                SpacingAfter = 8,
                WidowControl = true
            },
            HeadingStyles = new PdfHeadingStyles {
                Level1 = new PdfHeadingStyle {
                    FontSize = 20,
                    LineHeight = 1.15,
                    SpacingBefore = 0,
                    SpacingAfter = 8,
                    Color = headingText,
                    KeepWithNext = true
                },
                Level2 = new PdfHeadingStyle {
                    FontSize = 16,
                    LineHeight = 1.15,
                    SpacingBefore = 12,
                    SpacingAfter = 6,
                    Color = headingText,
                    KeepWithNext = true
                },
                Level3 = new PdfHeadingStyle {
                    FontSize = 13.5,
                    LineHeight = 1.15,
                    SpacingBefore = 8,
                    SpacingAfter = 4,
                    Color = headingText,
                    KeepWithNext = true
                }
            },
            ListStyle = new PdfListStyle {
                FontSize = 11,
                LineHeight = 1.15,
                LeftIndent = 18,
                MarkerGap = 6,
                ItemSpacing = 2,
                SpacingAfter = 8,
                Color = bodyText
            },
            TableStyle = tableStyle,
            PanelStyle = new PanelStyle {
                Background = PdfColor.FromRgb(248, 250, 252),
                BorderColor = ruleColor,
                BorderWidth = 0.6,
                PaddingX = 8,
                PaddingY = 6,
                SpacingBefore = 2,
                SpacingAfter = 8,
                KeepTogether = true
            },
            HorizontalRuleStyle = new PdfHorizontalRuleStyle {
                Thickness = 0.7,
                Color = ruleColor,
                SpacingBefore = 6,
                SpacingAfter = 8
            },
            ImageStyle = new PdfImageStyle {
                SpacingAfter = 8
            },
            DrawingStyle = new PdfDrawingStyle {
                SpacingAfter = 8
            },
            RowStyle = new PdfRowStyle {
                Gap = 18,
                SpacingAfter = 8
            }
        };
    }

    /// <summary>Creates a deep copy of this theme.</summary>
    public PdfTheme Clone() {
        return new PdfTheme {
            TextStyle = _textStyle?.Clone(),
            ParagraphStyle = _paragraphStyle?.Clone(),
            TableStyle = _tableStyle?.Clone(),
            HeadingStyles = _headingStyles?.Clone(),
            ListStyle = _listStyle?.Clone(),
            PanelStyle = _panelStyle?.Clone(),
            HorizontalRuleStyle = _horizontalRuleStyle?.Clone(),
            ImageStyle = _imageStyle?.Clone(),
            DrawingStyle = _drawingStyle?.Clone(),
            RowStyle = _rowStyle?.Clone()
        };
    }

    internal void ApplyTo(PdfOptions options) {
        Guard.NotNull(options, nameof(options));
        _textStyle?.ApplyTo(options);
        if (_paragraphStyle != null) {
            options.DefaultParagraphStyle = _paragraphStyle;
        }

        if (_tableStyle != null) {
            options.DefaultTableStyle = _tableStyle;
        }

        if (_headingStyles != null) {
            options.DefaultHeadingStyles = _headingStyles;
        }

        if (_listStyle != null) {
            options.DefaultListStyle = _listStyle;
        }

        if (_panelStyle != null) {
            options.DefaultPanelStyle = _panelStyle;
        }

        if (_horizontalRuleStyle != null) {
            options.DefaultHorizontalRuleStyle = _horizontalRuleStyle;
        }

        if (_imageStyle != null) {
            options.DefaultImageStyle = _imageStyle;
        }

        if (_drawingStyle != null) {
            options.DefaultDrawingStyle = _drawingStyle;
        }

        if (_rowStyle != null) {
            options.DefaultRowStyle = _rowStyle;
        }
    }
}
