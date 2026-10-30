namespace OfficeIMO.Pdf;

/// <summary>
/// Reusable bundle of default PDF styles for document or page composition.
/// </summary>
public sealed partial class PdfTheme {
    private PdfTextStyle? _textStyle;
    private PdfParagraphStyle? _paragraphStyle;
    private PdfTableStyle? _tableStyle;
    private PdfHeadingStyles? _headingStyles;
    private PdfListStyle? _listStyle;
    private PdfPanelStyle? _panelStyle;
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
    public PdfPanelStyle? PdfPanelStyle {
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
        global::OfficeIMO.OfficeColor bodyText = global::OfficeIMO.OfficeColor.FromRgb(31, 41, 55);
        global::OfficeIMO.OfficeColor headingText = global::OfficeIMO.OfficeColor.FromRgb(17, 24, 39);
        global::OfficeIMO.OfficeColor ruleColor = global::OfficeIMO.OfficeColor.FromRgb(203, 213, 225);

        var tableStyle = PdfTableStyles.ListTable1Light();
        tableStyle.TextColor = bodyText;
        tableStyle.HeaderTextColor = headingText;
        tableStyle.CellPaddingX = global::OfficeIMO.OfficeLength.Points(5);
        tableStyle.CellPaddingY = global::OfficeIMO.OfficeLength.Points(5);
        tableStyle.RowSeparatorColor = global::OfficeIMO.OfficeColor.FromRgb(226, 232, 240);
        tableStyle.HeaderSeparatorColor = headingText;
        tableStyle.FooterSeparatorColor = headingText;
        tableStyle.FooterSeparatorWidth = global::OfficeIMO.OfficeLength.Points(0.8);

        return new PdfTheme {
            TextStyle = new PdfTextStyle {
                Font = PdfStandardFont.Helvetica,
                FontSize = global::OfficeIMO.OfficeLength.Points(11),
                Color = bodyText
            },
            ParagraphStyle = new PdfParagraphStyle {
                LineHeight = 1.15,
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8),
                WidowControl = true
            },
            HeadingStyles = new PdfHeadingStyles {
                Level1 = new PdfHeadingStyle {
                    FontSize = global::OfficeIMO.OfficeLength.Points(20),
                    LineHeight = 1.15,
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(0),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(8),
                    Color = headingText,
                    KeepWithNext = true
                },
                Level2 = new PdfHeadingStyle {
                    FontSize = global::OfficeIMO.OfficeLength.Points(16),
                    LineHeight = 1.15,
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(12),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(6),
                    Color = headingText,
                    KeepWithNext = true
                },
                Level3 = new PdfHeadingStyle {
                    FontSize = global::OfficeIMO.OfficeLength.Points(13.5),
                    LineHeight = 1.15,
                    SpacingBefore = global::OfficeIMO.OfficeLength.Points(8),
                    SpacingAfter = global::OfficeIMO.OfficeLength.Points(4),
                    Color = headingText,
                    KeepWithNext = true
                }
            },
            ListStyle = new PdfListStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(11),
                LineHeight = 1.15,
                LeftIndent = global::OfficeIMO.OfficeLength.Points(18),
                MarkerGap = global::OfficeIMO.OfficeLength.Points(6),
                ItemSpacing = global::OfficeIMO.OfficeLength.Points(2),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8),
                Color = bodyText
            },
            TableStyle = tableStyle,
            PdfPanelStyle = new PdfPanelStyle {
                Background = global::OfficeIMO.OfficeColor.FromRgb(248, 250, 252),
                BorderColor = ruleColor,
                BorderWidth = global::OfficeIMO.OfficeLength.Points(0.6),
                PaddingX = global::OfficeIMO.OfficeLength.Points(8),
                PaddingY = global::OfficeIMO.OfficeLength.Points(6),
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(2),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8),
                KeepTogether = true
            },
            HorizontalRuleStyle = new PdfHorizontalRuleStyle {
                Thickness = global::OfficeIMO.OfficeLength.Points(0.7),
                Color = ruleColor,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(6),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
            },
            ImageStyle = new PdfImageStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
            },
            DrawingStyle = new PdfDrawingStyle {
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
            },
            RowStyle = new PdfRowStyle {
                Gap = global::OfficeIMO.OfficeLength.Points(18),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(8)
            }
        };
    }

    /// <summary>
    /// Creates a technical-document theme for specifications, READMEs, and operational guides with stronger heading contrast and compact table rhythm.
    /// </summary>
    public static PdfTheme TechnicalDocument() {
        PdfTheme theme = WordLike();
        theme.TableStyle = PdfTableStyles.TechnicalDocument();
        theme.PdfPanelStyle = new PdfPanelStyle {
            Background = Color(248, 250, 252),
            BorderColor = Color(203, 213, 225),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.5),
            PaddingX = global::OfficeIMO.OfficeLength.Points(9),
            PaddingY = global::OfficeIMO.OfficeLength.Points(7),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(2),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(9),
            KeepTogether = true
        };
        theme.HorizontalRuleStyle = new PdfHorizontalRuleStyle {
            Color = Color(203, 213, 225),
            Thickness = global::OfficeIMO.OfficeLength.Points(0.6),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(8),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(10)
        };
        return theme;
    }

    /// <summary>
    /// Creates a compact document theme for dense notes and narrow reports that still need readable rhythm.
    /// </summary>
    public static PdfTheme Compact() {
        PdfTheme theme = WordLike();
        global::OfficeIMO.OfficeColor bodyText = Color(31, 41, 55);
        theme.TextStyle = new PdfTextStyle {
            Font = PdfStandardFont.Helvetica,
            FontSize = global::OfficeIMO.OfficeLength.Points(10),
            Color = bodyText
        };
        theme.ParagraphStyle = new PdfParagraphStyle {
            LineHeight = 1.08,
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(4),
            WidowControl = true
        };
        theme.ListStyle = new PdfListStyle {
            FontSize = global::OfficeIMO.OfficeLength.Points(10),
            LineHeight = 1.08,
            LeftIndent = global::OfficeIMO.OfficeLength.Points(15),
            MarkerGap = global::OfficeIMO.OfficeLength.Points(5),
            ItemSpacing = global::OfficeIMO.OfficeLength.Points(1),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(4),
            Color = bodyText
        };
        theme.TableStyle = PdfTableStyles.Compact();
        theme.PdfPanelStyle = new PdfPanelStyle {
            Background = Color(248, 250, 252),
            BorderColor = Color(203, 213, 225),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.4),
            PaddingX = global::OfficeIMO.OfficeLength.Points(7),
            PaddingY = global::OfficeIMO.OfficeLength.Points(5),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(1),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(5),
            KeepTogether = true
        };
        theme.HorizontalRuleStyle = new PdfHorizontalRuleStyle {
            Color = Color(226, 232, 240),
            Thickness = global::OfficeIMO.OfficeLength.Points(0.5),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(5),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(6)
        };
        theme.RowStyle = new PdfRowStyle {
            Gap = global::OfficeIMO.OfficeLength.Points(14),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(5)
        };
        return theme;
    }

    /// <summary>
    /// Creates a report-oriented theme with stronger table hierarchy, summary-row separators, and calm document colors.
    /// </summary>
    public static PdfTheme Report() {
        PdfTheme theme = WordLike();
        global::OfficeIMO.OfficeColor headingText = Color(30, 41, 59);
        global::OfficeIMO.OfficeColor accent = Color(30, 64, 175);
        theme.TextStyle = new PdfTextStyle {
            Font = PdfStandardFont.Helvetica,
            FontSize = global::OfficeIMO.OfficeLength.Points(10.5),
            Color = headingText
        };
        theme.HeadingStyles = new PdfHeadingStyles {
            Level1 = new PdfHeadingStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(21),
                LineHeight = 1.12,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(0),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(9),
                Color = headingText,
                KeepWithNext = true
            },
            Level2 = new PdfHeadingStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(15.5),
                LineHeight = 1.14,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(12),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(6),
                Color = accent,
                KeepWithNext = true
            },
            Level3 = new PdfHeadingStyle {
                FontSize = global::OfficeIMO.OfficeLength.Points(12.75),
                LineHeight = 1.15,
                SpacingBefore = global::OfficeIMO.OfficeLength.Points(8),
                SpacingAfter = global::OfficeIMO.OfficeLength.Points(4),
                Color = headingText,
                KeepWithNext = true
            }
        };
        theme.TableStyle = PdfTableStyles.Report();
        theme.PdfPanelStyle = new PdfPanelStyle {
            Background = Color(239, 246, 255),
            BorderColor = Color(191, 219, 254),
            BorderWidth = global::OfficeIMO.OfficeLength.Points(0.6),
            PaddingX = global::OfficeIMO.OfficeLength.Points(10),
            PaddingY = global::OfficeIMO.OfficeLength.Points(8),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(3),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(10),
            KeepTogether = true
        };
        theme.HorizontalRuleStyle = new PdfHorizontalRuleStyle {
            Color = Color(191, 219, 254),
            Thickness = global::OfficeIMO.OfficeLength.Points(0.75),
            SpacingBefore = global::OfficeIMO.OfficeLength.Points(8),
            SpacingAfter = global::OfficeIMO.OfficeLength.Points(10)
        };
        return theme;
    }

    /// <summary>Creates a deep copy of this theme.</summary>
    public PdfTheme Clone() {
        return new PdfTheme {
            TextStyle = _textStyle?.Clone(),
            ParagraphStyle = _paragraphStyle?.Clone(),
            TableStyle = _tableStyle?.Clone(),
            HeadingStyles = _headingStyles?.Clone(),
            ListStyle = _listStyle?.Clone(),
            PdfPanelStyle = _panelStyle?.Clone(),
            HorizontalRuleStyle = _horizontalRuleStyle?.Clone(),
            ImageStyle = _imageStyle?.Clone(),
            DrawingStyle = _drawingStyle?.Clone(),
            RowStyle = _rowStyle?.Clone()
        };
    }

    internal PdfTableStyle? TableStyleSnapshot => _tableStyle;
    internal PdfHeadingStyles? HeadingStylesSnapshot => _headingStyles;
    internal PdfListStyle? ListStyleSnapshot => _listStyle;
    internal PdfPanelStyle? PanelStyleSnapshot => _panelStyle;
    internal PdfParagraphStyle? ParagraphStyleSnapshot => _paragraphStyle;
    internal PdfHorizontalRuleStyle? HorizontalRuleStyleSnapshot => _horizontalRuleStyle;

    internal void ApplyTo(PdfCreateOptions options) {
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

    private static global::OfficeIMO.OfficeColor Color(byte red, byte green, byte blue) => global::OfficeIMO.OfficeColor.FromRgb(red, green, blue);
}
