using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfTheme {
    private const double HeadingLineHeight = 1.15D;
    private const double ListItemSpacingRatio = 0.5D;

    /// <summary>
    /// Creates PDF composition defaults from a cross-format <see cref="OfficeTheme"/>: the type scale and heading
    /// colour for headings, the theme table style with an accent header rule and hairline rows, list indentation with
    /// accent bullets, and paragraph spacing from the theme rhythm.
    /// </summary>
    /// <param name="theme">Theme that supplies colours, sizes and spacing.</param>
    public static PdfTheme FromOfficeTheme(OfficeTheme theme) {
        Guard.NotNull(theme, nameof(theme));
        OfficeColor heading = Resolve(theme.TextStyle.HeadingColor, theme);
        OfficeColor accent = Resolve(theme.Palette.Accent1, theme);
        double paragraphAfter = theme.Spacing.ParagraphAfter.ToPoints();

        return new PdfTheme {
            TextStyle = new PdfTextStyle {
                Font = ResolveStandardFont(theme.Fonts.Body.Family),
                FontSize = theme.TypeScale.Body.ToPoints(),
                Color = Resolve(theme.Palette.Text, theme)
            },
            ParagraphStyle = new PdfParagraphStyle {
                LineHeight = theme.TextStyle.LineSpacing,
                SpacingBefore = theme.Spacing.ParagraphBefore.ToPoints(),
                SpacingAfter = paragraphAfter,
                WidowControl = true
            },
            HeadingStyles = new PdfHeadingStyles {
                Level1 = CreateHeading(theme, 1, theme.TypeScale.Heading1.ToPoints(), heading),
                Level2 = CreateHeading(theme, 2, theme.TypeScale.Heading2.ToPoints(), heading),
                Level3 = CreateHeading(theme, 3, theme.TypeScale.Heading3.ToPoints(), heading)
            },
            ListStyle = new PdfListStyle {
                LeftIndent = theme.Spacing.ListIndent.ToPoints(),
                MarkerGap = theme.TypeScale.Body.ToPoints() * 0.5D,
                MarkerColor = accent,
                ItemSpacing = paragraphAfter * ListItemSpacingRatio,
                SpacingAfter = paragraphAfter
            },
            TableStyle = CreateTableStyle(theme, paragraphAfter),
            PdfPanelStyle = new PdfPanelStyle {
                Background = Resolve(theme.Surfaces.Surface, theme),
                BorderColor = Resolve(theme.Surfaces.Rule, theme),
                BorderWidth = theme.TableStyle.BorderWidth.ToPoints(),
                PaddingX = theme.TableStyle.CellPadding.ToPoints() * 2D,
                PaddingY = theme.TableStyle.CellPadding.ToPoints() * 1.4D,
                SpacingBefore = theme.Spacing.ParagraphBefore.ToPoints(),
                SpacingAfter = paragraphAfter,
                KeepTogether = true
            },
            HorizontalRuleStyle = new PdfHorizontalRuleStyle {
                Color = Resolve(theme.Surfaces.Rule, theme),
                Thickness = theme.TableStyle.BorderWidth.ToPoints(),
                SpacingBefore = paragraphAfter * 0.5D,
                SpacingAfter = paragraphAfter
            }
        };
    }

    internal static PdfTableStyle CreateTableStyle(OfficeTheme theme, double spacingAfter) {
        double padding = theme.TableStyle.CellPadding.ToPoints();
        OfficeColor border = Resolve(theme.TableStyle.BorderColor, theme);
        OfficeColor text = Resolve(theme.Palette.Text, theme);
        OfficeThemeRule headerRule = theme.TableStyle.HeaderRule;
        return new PdfTableStyle {
            HeaderFill = Resolve(theme.TableStyle.HeaderFill, theme),
            HeaderTextColor = Resolve(theme.TableStyle.HeaderText, theme),
            HeaderBold = theme.TableStyle.HeaderBold,
            TextColor = text,
            RowStripeFill = theme.TableStyle.Banding ? Resolve(theme.TableStyle.BandFill, theme) : null,
            // Horizontal hairlines only, with the theme's rule under the header row.
            BorderColor = null,
            BorderWidth = 0D,
            RowSeparatorColor = border,
            RowSeparatorWidth = theme.TableStyle.BorderWidth.ToPoints(),
            HeaderSeparatorColor = headerRule.IsVisible ? Resolve(headerRule.Color, theme) : null,
            HeaderSeparatorWidth = headerRule.IsVisible ? headerRule.Width.ToPoints() : 0D,
            FooterSeparatorColor = text,
            FooterSeparatorWidth = headerRule.IsVisible ? headerRule.Width.ToPoints() : theme.TableStyle.BorderWidth.ToPoints(),
            CellPaddingTop = padding,
            CellPaddingBottom = padding,
            CellPaddingX = padding * 1.4D,
            FontSize = System.Math.Max(1D, theme.TypeScale.Body.ToPoints() - 1D),
            HeaderFontSize = System.Math.Max(1D, theme.TypeScale.Body.ToPoints() - 1D),
            SpacingBefore = theme.Spacing.ParagraphBefore.ToPoints(),
            SpacingAfter = spacingAfter,
            AutoFitColumns = true,
            RightAlignNumeric = true,
            CaptionColor = Resolve(theme.Palette.TextMuted, theme),
            CaptionFontSize = theme.TypeScale.Caption.ToPoints()
        };
    }

    private static PdfHeadingStyle CreateHeading(OfficeTheme theme, int level, double size, OfficeColor color) {
        OfficeThemeParagraphSpacing spacing = theme.TextStyle.GetHeadingSpacing(level);
        var style = new PdfHeadingStyle {
            FontSize = size,
            LineHeight = HeadingLineHeight,
            SpacingBefore = spacing.Before.ToPoints(),
            SpacingAfter = spacing.After.ToPoints(),
            Bold = theme.TextStyle.HeadingBold,
            Color = color,
            KeepWithNext = true
        };
        ApplyThemeHeadingFont(theme, style);
        return style;
    }

    internal static void ApplyThemeHeadingFont(OfficeTheme theme, PdfHeadingStyle style) {
        if (string.Equals(theme.Fonts.Heading.Family, theme.Fonts.Body.Family, StringComparison.OrdinalIgnoreCase)) return;
        style.FontFamily = theme.Fonts.Heading.Family;
        style.Font = ResolveStandardFont(theme.Fonts.Heading.Family);
    }

    // Embedding remains with the existing explicit PDF font APIs.
    internal static PdfStandardFont ResolveStandardFont(string family) =>
        OfficeDocumentFontSubstitutions.ClassifyGenericFamily(family) switch {
            "serif" => PdfStandardFont.TimesRoman,
            "monospace" => PdfStandardFont.Courier,
            _ => PdfStandardFont.Helvetica
        };

    internal static OfficeColor Resolve(OfficeColor color, OfficeTheme theme) =>
        color.IsThemeRole ? color.Resolve(theme) : color;
}
