using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfTheme {
    // Values the visual design needs that OfficeTheme does not expose yet.
    // They are derived from theme values rather than copied from themes.json.
    private const double HeadingSpacingBeforeRatio = 0.9D;
    private const double HeadingSpacingAfterRatio = 0.3D;
    private const double HeaderRuleWidth = 1D;
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
        PdfColor text = Resolve(theme.Palette.Text, theme);
        PdfColor accent = Resolve(theme.Palette.Accent1, theme);
        double paragraphAfter = theme.Spacing.ParagraphAfter.ToPoints();

        return new PdfTheme {
            HeadingStyles = new PdfHeadingStyles {
                Level1 = CreateHeading(theme.TypeScale.Heading1.ToPoints(), text),
                Level2 = CreateHeading(theme.TypeScale.Heading2.ToPoints(), text),
                Level3 = CreateHeading(theme.TypeScale.Heading3.ToPoints(), text)
            },
            ListStyle = new PdfListStyle {
                LeftIndent = theme.Spacing.ListIndent.ToPoints(),
                MarkerColor = accent,
                ItemSpacing = paragraphAfter * ListItemSpacingRatio,
                SpacingAfter = paragraphAfter
            },
            TableStyle = CreateTableStyle(theme, paragraphAfter),
            PdfPanelStyle = new PdfPanelStyle {
                Background = Resolve(theme.TableStyle.BandFill, theme),
                BorderColor = Resolve(theme.TableStyle.BorderColor, theme),
                BorderWidth = theme.TableStyle.BorderWidth.ToPoints(),
                PaddingX = theme.TableStyle.CellPadding.ToPoints() * 2D,
                PaddingY = theme.TableStyle.CellPadding.ToPoints() * 1.4D,
                SpacingBefore = theme.Spacing.ParagraphBefore.ToPoints(),
                SpacingAfter = paragraphAfter,
                KeepTogether = true
            },
            HorizontalRuleStyle = new PdfHorizontalRuleStyle {
                Color = Resolve(theme.TableStyle.BorderColor, theme),
                Thickness = theme.TableStyle.BorderWidth.ToPoints(),
                SpacingBefore = paragraphAfter * 0.5D,
                SpacingAfter = paragraphAfter
            }
        };
    }

    internal static PdfTableStyle CreateTableStyle(OfficeTheme theme, double spacingAfter) {
        double padding = theme.TableStyle.CellPadding.ToPoints();
        PdfColor border = Resolve(theme.TableStyle.BorderColor, theme);
        PdfColor text = Resolve(theme.Palette.Text, theme);
        return new PdfTableStyle {
            HeaderFill = Resolve(theme.TableStyle.HeaderFill, theme),
            HeaderTextColor = Resolve(theme.TableStyle.HeaderText, theme),
            HeaderBold = true,
            TextColor = text,
            RowStripeFill = theme.TableStyle.Banding ? Resolve(theme.TableStyle.BandFill, theme) : null,
            // Horizontal hairlines only, with the accent rule under the header row.
            BorderColor = null,
            BorderWidth = 0D,
            RowSeparatorColor = border,
            RowSeparatorWidth = theme.TableStyle.BorderWidth.ToPoints(),
            HeaderSeparatorColor = Resolve(theme.Palette.Accent1, theme),
            HeaderSeparatorWidth = HeaderRuleWidth,
            FooterSeparatorColor = text,
            FooterSeparatorWidth = HeaderRuleWidth,
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

    private static PdfHeadingStyle CreateHeading(double size, PdfColor color) => new PdfHeadingStyle {
        FontSize = size,
        LineHeight = HeadingLineHeight,
        SpacingBefore = System.Math.Round(size * HeadingSpacingBeforeRatio, 1),
        SpacingAfter = System.Math.Round(size * HeadingSpacingAfterRatio, 1),
        Color = color,
        KeepWithNext = true
    };

    private static PdfColor Resolve(OfficeColor color, OfficeTheme theme) =>
        PdfColor.FromOfficeColor(color.IsThemeRole ? color.Resolve(theme) : color);
}
