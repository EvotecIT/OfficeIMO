using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Markdown.Pdf;

public sealed partial class MarkdownPdfStyle {
    /// <summary>Creates a PDF-specific profile from a shared Markdown visual theme.</summary>
    internal static MarkdownPdfStyle FromMarkdownTheme(MarkdownVisualTheme theme) {
        if (theme == null) {
            throw new ArgumentNullException(nameof(theme));
        }

        MarkdownPdfStyle pdfTheme = Create(theme.Kind);
        pdfTheme.Name = theme.Name;
        ApplySharedPalette(pdfTheme, theme.PaletteSnapshot, theme.TableSnapshot);
        return pdfTheme;
    }

    private static void ApplySharedPalette(MarkdownPdfStyle pdfTheme, MarkdownVisualPalette palette, MarkdownTableVisualStyle table) {
        PdfCore.PdfColor accent = ToPdfColorOrNull(palette.Accent) ?? pdfTheme.LinkColorSnapshot;
        PdfCore.PdfColor heading = ToPdfColorOrNull(palette.Heading) ?? pdfTheme.DocumentHeaderTitleColorSnapshot;
        PdfCore.PdfColor text = ToPdfColorOrNull(palette.Text) ?? pdfTheme.CodeBlockTextColorSnapshot;
        PdfCore.PdfColor muted = ToPdfColorOrNull(palette.MutedText) ?? pdfTheme.DocumentHeaderMetadataColorSnapshot;
        PdfCore.PdfColor? surface = ToPdfColorOrNull(palette.Surface);
        PdfCore.PdfColor? border = ToPdfColorOrNull(palette.Border);

        ApplySharedDocumentTheme(pdfTheme, heading, text);
        ApplySharedPageBackground(pdfTheme, palette.Background);
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
        pdfTheme.ChecklistCheckedIconColor = ToPdfColor(OfficeColor.Parse("SeaGreen"));
        pdfTheme.ChecklistUncheckedIconColor = muted;
        pdfTheme.ChecklistCheckedTextColor = muted;
        pdfTheme.ChecklistUncheckedTextColor = text;
        pdfTheme.ChecklistCheckedFillColor = surface;

        PdfCore.PdfTableStyle tableStyle = pdfTheme.TableStyleSnapshot;
        tableStyle.BorderColor = border;
        tableStyle.RowSeparatorColor = border;
        tableStyle.HeaderFill = table.EmphasizeHeader ? ToPdfColorOrNull(palette.TableHeaderBackground) : null;
        tableStyle.HeaderTextColor = table.EmphasizeHeader ? ToPdfColorOrNull(palette.TableHeaderText) : null;
        tableStyle.TextColor = text;
        tableStyle.RowStripeFill = table.UseRowStripes ? ToPdfColorOrNull(palette.TableStripeBackground) : null;
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
        MarkdownPdfFigureStyle figureStyle = MarkdownPdfFigureStyle.Framed(surface ?? PdfCore.PdfColor.White, border ?? PdfCore.PdfColor.LightGray, accent, muted, table.BorderWidth);
        if (surface == null && figureStyle.PanelStyle != null) {
            PdfCore.PanelStyle panelStyle = figureStyle.PanelStyle;
            panelStyle.Background = null;
            figureStyle.PanelStyle = panelStyle;
        }

        pdfTheme.FigureStyle = figureStyle;
    }

    private static void ApplySharedPageBackground(MarkdownPdfStyle pdfTheme, OfficeColor background) {
        if (background.A == 0) {
            pdfTheme.PageDecoration = null;
            return;
        }

        MarkdownPdfPageDecoration decoration = pdfTheme.PageDecoration ?? new MarkdownPdfPageDecoration();
        decoration.BackgroundColor = ToPdfColor(background);
        pdfTheme.PageDecoration = decoration;
    }

    private static void ApplySharedDocumentTheme(MarkdownPdfStyle pdfTheme, PdfCore.PdfColor heading, PdfCore.PdfColor text) {
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

    private static PdfCore.PanelStyle PanelFromShared(OfficeColor background, OfficeColor border, double borderWidth, double paddingX, double paddingY, double spacingBefore, double spacingAfter) => new PdfCore.PanelStyle {
        Background = ToPdfColorOrNull(background),
        BorderColor = ToPdfColorOrNull(border),
        BorderWidth = borderWidth,
        PaddingX = paddingX,
        PaddingY = paddingY,
        SpacingBefore = spacingBefore,
        SpacingAfter = spacingAfter,
        KeepTogether = true
    };

    private static PdfCore.PanelStyle PanelWithLeftRuleFromShared(OfficeColor background, OfficeColor border, OfficeColor left, double borderWidth, double leftWidth, double paddingX, double paddingY, double spacingBefore, double spacingAfter) {
        PdfCore.PanelStyle style = PanelFromShared(background, border, borderWidth, paddingX, paddingY, spacingBefore, spacingAfter);
        style.LeftBorder = new PdfCore.PdfPanelBorder {
            Color = ToPdfColorOrNull(left),
            Width = leftWidth
        };
        return style;
    }

    private static PdfCore.PdfColor ToPdfColor(OfficeColor color) => PdfCore.PdfColor.FromRgb(color.R, color.G, color.B);

    private static PdfCore.PdfColor? ToPdfColorOrNull(OfficeColor color) => color.A == 0 ? null : ToPdfColor(color);
}
