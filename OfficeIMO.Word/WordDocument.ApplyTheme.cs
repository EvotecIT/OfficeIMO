using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>
    /// Restyles the current document with the shared theme while retaining content, semantic emphasis,
    /// fields, relationships, table sizing and section geometry. Typography and cell padding can change pagination.
    /// </summary>
    /// <param name="theme">The visual theme to apply explicitly.</param>
    /// <returns>The current editable document.</returns>
    public WordDocument ApplyTheme(OfficeTheme theme) {
        ArgumentNullException.ThrowIfNull(theme);
        if (_disposed) throw new ObjectDisposedException(nameof(WordDocument));
        if (FileOpenAccess == FileAccess.Read) throw new InvalidOperationException("A read-only document cannot be themed.");
        MainDocumentPart main = _wordprocessingDocument.MainDocumentPart!;
        Styles? sourceStyles = main.StyleDefinitionsPart?.Styles;
        OpenXmlCompositeElement[] roots = WordFieldInventory.EnumerateFieldRoots(main).Select(story => story.Root)
            .Concat(main.WordprocessingCommentsPart?.Comments is { } comments
                ? new OpenXmlCompositeElement[] { comments } : Array.Empty<OpenXmlCompositeElement>()).ToArray();
        var appearances = new Dictionary<ThemeTextRole, Style>();

        using IDisposable formattingScope = WordRunFormatting.BeginConversionScope();
        // Resolve authored roles/fonts before changing the shared defaults or native font scheme.
        foreach (OpenXmlCompositeElement root in roots) {
            foreach (Paragraph paragraph in root.Descendants<Paragraph>()) {
                if (paragraph.Ancestors<Footnote>().Any(note => note.Type?.Value is { } type && type != FootnoteEndnoteValues.Normal) ||
                    paragraph.Ancestors<Endnote>().Any(note => note.Type?.Value is { } type && type != FootnoteEndnoteValues.Normal)) continue;
                ThemeTextRole role = ResolveThemeRole(paragraph, sourceStyles, root is Footer);
                if (!appearances.TryGetValue(role, out Style? appearance)) {
                    appearance = CreateThemeAppearance(role, theme);
                    appearances.Add(role, appearance);
                }
                ApplyThemeParagraph(paragraph, role, appearance, theme);
                foreach (Run run in paragraph.Descendants<Run>().Where(run => ReferenceEquals(run.Ancestors<Paragraph>().FirstOrDefault(), paragraph))) {
                    if (!run.Elements<Text>().Any() && !run.Elements<SymbolChar>().Any() && !run.Elements<TabChar>().Any()) continue;
                    ApplyThemeRun(paragraph, run, role, appearance, sourceStyles, theme);
                }
            }
        }
        foreach (Table table in roots.SelectMany(root => root.Descendants<Table>())) ApplyThemeTable(table, theme);
        Styles themedStyles = sourceStyles != null ? (Styles)sourceStyles.CloneNode(true) : new Styles();
        WordThemeDefaults.ApplyCreationTypography(themedStyles, theme);
        ApplyThemeDefaults(themedStyles, theme, WordThemeDefaults.CreateBody(theme).StyleParagraphProperties!.SpacingBetweenLines);
        return this;
    }

    private enum ThemeTextRole { Body, Heading1, Heading2, Heading3, Heading4, Heading5, Heading6, Heading7, Heading8, Heading9, Title, Subtitle, Code, Table, TableHeader, Footer }

    private static ThemeTextRole ResolveThemeRole(Paragraph paragraph, Styles? styles, bool footer) {
        ThemeTextRole context = paragraph.Ancestors<TableCell>().FirstOrDefault() is { } cell
            ? IsThemeHeaderCell(cell) ? ThemeTextRole.TableHeader : ThemeTextRole.Table
            : footer ? ThemeTextRole.Footer : ThemeTextRole.Body;
        if (paragraph.ParagraphProperties?.OutlineLevel?.Val?.Value is int directOutline)
            return directOutline >= 0 && directOutline <= 8 ? (ThemeTextRole)(directOutline + 1) : context;
        string? styleId = paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        Style[] chain = WordRunFormatting.EnumerateStyleChain(styles, styleId, StyleValues.Paragraph).ToArray();
        if (chain.Length == 0) {
            if (styleId == "Title") return ThemeTextRole.Title;
            if (styleId == "Subtitle") return ThemeTextRole.Subtitle;
            if (styleId == "OfficeIMOCodeBlock") return ThemeTextRole.Code;
            if (WordThemeDefaults.TryHeadingLevel(styleId, out int level)) return (ThemeTextRole)level;
        }
        foreach (Style style in chain) {
            string? id = style.StyleId?.Value;
            if (id == "Title") return ThemeTextRole.Title;
            if (id == "Subtitle") return ThemeTextRole.Subtitle;
            if (id == "OfficeIMOCodeBlock") return ThemeTextRole.Code;
            if (style.StyleParagraphProperties?.OutlineLevel?.Val?.Value is int outline)
                return outline >= 0 && outline <= 8 ? (ThemeTextRole)(outline + 1) : context;
            if (WordThemeDefaults.TryHeadingLevel(id, out int level)) return (ThemeTextRole)level;
        }
        return context;
    }

    private static Style CreateThemeAppearance(ThemeTextRole role, OfficeTheme theme) {
        if (role >= ThemeTextRole.Heading1 && role <= ThemeTextRole.Heading9) return WordThemeDefaults.CreateHeading((int)role, theme);
        if (role == ThemeTextRole.Title) return WordThemeDefaults.CreateTitle(theme);
        if (role == ThemeTextRole.Subtitle) return WordThemeDefaults.CreateSubtitle(theme);
        if (role == ThemeTextRole.Table || role == ThemeTextRole.TableHeader)
            return WordThemeDefaults.CreateTableTextStyle(theme, role == ThemeTextRole.TableHeader);
        Style body = WordThemeDefaults.CreateBody(theme);
        if (role == ThemeTextRole.Footer) {
            body.StyleRunProperties!.Color = new Color { Val = theme.Footer.Color.Resolve(theme).ToRgbHex() };
            body.StyleRunProperties.FontSize = new FontSize { Val = WordThemeDefaults.HalfPoints(theme.Footer.FontSize) };
            body.StyleRunProperties.FontSizeComplexScript = new FontSizeComplexScript { Val = WordThemeDefaults.HalfPoints(theme.Footer.FontSize) };
        }
        return body;
    }
}
