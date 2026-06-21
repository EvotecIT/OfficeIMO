using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

public sealed partial class MarkdownPdfVisualTheme {
    /// <summary>Creates a PDF-specific profile from a shared Markdown visual theme.</summary>
    public static MarkdownPdfVisualTheme FromMarkdownTheme(MarkdownVisualTheme theme) {
        if (theme == null) {
            throw new ArgumentNullException(nameof(theme));
        }

        MarkdownPdfVisualTheme pdfTheme = Create(MapKind(theme.Kind));
        pdfTheme.Name = theme.Name;
        ApplySharedPalette(pdfTheme, theme.PaletteSnapshot, theme.TableSnapshot);
        return pdfTheme;
    }

    private static MarkdownPdfThemeKind MapKind(MarkdownVisualThemeKind kind) {
        switch (kind) {
            case MarkdownVisualThemeKind.Plain:
                return MarkdownPdfThemeKind.Plain;
            case MarkdownVisualThemeKind.WordLike:
                return MarkdownPdfThemeKind.WordLike;
            case MarkdownVisualThemeKind.TechnicalDocument:
                return MarkdownPdfThemeKind.TechnicalDocument;
            case MarkdownVisualThemeKind.GitHubLike:
                return MarkdownPdfThemeKind.GitHubLike;
            case MarkdownVisualThemeKind.Compact:
                return MarkdownPdfThemeKind.Compact;
            case MarkdownVisualThemeKind.Report:
                return MarkdownPdfThemeKind.Report;
            default:
                return MarkdownPdfThemeKind.WordLike;
        }
    }

    private static void ApplySharedPalette(MarkdownPdfVisualTheme pdfTheme, MarkdownVisualPalette palette, MarkdownTableVisualStyle table) {
        PdfCore.PdfColor accent = ToPdfColor(palette.Accent);
        PdfCore.PdfColor heading = ToPdfColor(palette.Heading);
        PdfCore.PdfColor text = ToPdfColor(palette.Text);
        PdfCore.PdfColor muted = ToPdfColor(palette.MutedText);
        PdfCore.PdfColor surface = ToPdfColor(palette.Surface);
        PdfCore.PdfColor border = ToPdfColor(palette.Border);

        ApplySharedDocumentTheme(pdfTheme, heading, text);
        pdfTheme.DocumentHeaderTitleColor = heading;
        pdfTheme.DocumentHeaderSubtitleColor = accent;
        pdfTheme.DocumentHeaderMetadataColor = muted;
        pdfTheme.DocumentHeaderRuleColor = accent;
        pdfTheme.TocTitleColor = heading;
        pdfTheme.TocLinkColor = accent;
        pdfTheme.TocTextColor = muted;
        pdfTheme.LinkColor = accent;
        pdfTheme.CodeBlockTextColor = text;
        pdfTheme.CodeBlockLabelColor = accent;
        pdfTheme.SemanticBlockTextColor = text;
        pdfTheme.SemanticBlockLabelColor = accent;
        pdfTheme.ChecklistCheckedIconColor = ToPdfColor(MarkdownColor.Parse("SeaGreen"));
        pdfTheme.ChecklistUncheckedIconColor = muted;
        pdfTheme.ChecklistCheckedTextColor = muted;
        pdfTheme.ChecklistUncheckedTextColor = text;
        pdfTheme.ChecklistCheckedFillColor = surface;

        PdfCore.PdfTableStyle tableStyle = pdfTheme.TableStyleSnapshot;
        tableStyle.BorderColor = border;
        tableStyle.RowSeparatorColor = border;
        tableStyle.HeaderFill = ToPdfColor(palette.TableHeaderBackground);
        tableStyle.HeaderTextColor = ToPdfColor(palette.TableHeaderText);
        tableStyle.TextColor = text;
        tableStyle.RowStripeFill = table.UseRowStripes ? ToPdfColor(palette.TableStripeBackground) : null;
        tableStyle.BorderWidth = table.BorderWidth;
        tableStyle.CellPaddingX = table.CellPaddingX;
        tableStyle.CellPaddingY = table.CellPaddingY;
        pdfTheme.TableStyle = tableStyle;

        PdfCore.PdfTableStyle frontMatterStyle = pdfTheme.FrontMatterTableStyleSnapshot;
        frontMatterStyle.BorderColor = border;
        frontMatterStyle.RowSeparatorColor = border;
        frontMatterStyle.HeaderFill = surface;
        frontMatterStyle.HeaderTextColor = heading;
        frontMatterStyle.TextColor = text;
        pdfTheme.FrontMatterTableStyle = frontMatterStyle;

        pdfTheme.CodeBlockPanelStyle = PanelFromShared(palette.CodeBackground, palette.Border, table.BorderWidth, 9, 7, 4, 8);
        pdfTheme.SemanticBlockPanelStyle = PanelFromShared(palette.Surface, palette.Accent, table.BorderWidth, 9, 7, 4, 8);
        pdfTheme.QuotePanelStyle = PanelWithLeftRuleFromShared(palette.Background, palette.Border, palette.Accent, 0.0, 3, 10, 8, 2, 9);
        pdfTheme.CalloutPanelStyle = PanelFromShared(palette.Surface, palette.Accent, Math.Max(0.7, table.BorderWidth), 10, 8, 4, 9);
        pdfTheme.TocPanelStyle = PanelWithLeftRuleFromShared(palette.Surface, palette.Border, palette.Accent, table.BorderWidth, 3, 11, 8, 4, 10);
        pdfTheme.FigureStyle = MarkdownPdfFigureStyle.Framed(surface, border, accent, muted, table.BorderWidth);
    }

    private static void ApplySharedDocumentTheme(MarkdownPdfVisualTheme pdfTheme, PdfCore.PdfColor heading, PdfCore.PdfColor text) {
        PdfCore.PdfTheme documentTheme = pdfTheme.DocumentThemeSnapshot ?? PdfCore.PdfTheme.WordLike();
        PdfCore.PdfTextStyle textStyle = documentTheme.TextStyle ?? new PdfCore.PdfTextStyle();
        textStyle.Color = text;
        documentTheme.TextStyle = textStyle;

        PdfCore.PdfHeadingStyles headingStyles = documentTheme.HeadingStyles ?? new PdfCore.PdfHeadingStyles();
        ApplyHeadingColor(headingStyles, 1, heading);
        ApplyHeadingColor(headingStyles, 2, heading);
        ApplyHeadingColor(headingStyles, 3, heading);
        documentTheme.HeadingStyles = headingStyles;
        pdfTheme.DocumentTheme = documentTheme;
    }

    private static void ApplyHeadingColor(PdfCore.PdfHeadingStyles headingStyles, int level, PdfCore.PdfColor color) {
        PdfCore.PdfHeadingStyle style = level switch {
            1 => headingStyles.Level1 ?? new PdfCore.PdfHeadingStyle(),
            2 => headingStyles.Level2 ?? new PdfCore.PdfHeadingStyle(),
            _ => headingStyles.Level3 ?? new PdfCore.PdfHeadingStyle()
        };
        style.Color = color;
        switch (level) {
            case 1:
                headingStyles.Level1 = style;
                break;
            case 2:
                headingStyles.Level2 = style;
                break;
            default:
                headingStyles.Level3 = style;
                break;
        }
    }

    private static PdfCore.PanelStyle PanelFromShared(MarkdownColor background, MarkdownColor border, double borderWidth, double paddingX, double paddingY, double spacingBefore, double spacingAfter) => new PdfCore.PanelStyle {
        Background = ToPdfColor(background),
        BorderColor = ToPdfColor(border),
        BorderWidth = borderWidth,
        PaddingX = paddingX,
        PaddingY = paddingY,
        SpacingBefore = spacingBefore,
        SpacingAfter = spacingAfter,
        KeepTogether = true
    };

    private static PdfCore.PanelStyle PanelWithLeftRuleFromShared(MarkdownColor background, MarkdownColor border, MarkdownColor left, double borderWidth, double leftWidth, double paddingX, double paddingY, double spacingBefore, double spacingAfter) {
        PdfCore.PanelStyle style = PanelFromShared(background, border, borderWidth, paddingX, paddingY, spacingBefore, spacingAfter);
        style.LeftBorder = new PdfCore.PdfPanelBorder {
            Color = ToPdfColor(left),
            Width = leftWidth
        };
        return style;
    }

    private static PdfCore.PdfColor ToPdfColor(MarkdownColor color) => PdfCore.PdfColor.FromRgb(color.R, color.G, color.B);
}
