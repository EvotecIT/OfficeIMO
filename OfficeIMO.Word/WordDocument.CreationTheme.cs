using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private OfficeTheme? _creationTheme;
    // Shared theme components are immutable. Ordinary loads never enter the creation-theme path.
    internal OfficeTheme EffectiveCreationTheme => _creationTheme ?? OfficeTheme.Default;

    private static WordDocument ApplyTemplateCreationTheme(WordDocument document, OfficeTheme? theme) {
        if (theme == null) return document;
        try {
            document.ApplyCreationTheme(theme);
            return document;
        } catch {
            document.Dispose();
            throw;
        }
    }

    /// <summary>Projects an explicitly selected creation theme into existing role definitions and native defaults.</summary>
    private void ApplyCreationTheme(OfficeTheme theme) {
        MainDocumentPart main = _wordprocessingDocument.MainDocumentPart!;
        StyleDefinitionsPart part = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        // Clone the package's catalog, never the static shared catalog's caller-specific contents.
        Styles styles = part.Styles != null ? (Styles)part.Styles.CloneNode(true) : new Styles();
        ReplaceCreationTableStyle(styles, theme);
        WordThemeDefaults.ApplyCreationTypography(styles, theme);
        ApplyThemeDefaults(styles, theme, WordThemeDefaults.CreateBody(theme).StyleParagraphProperties!.SpacingBetweenLines);
    }

    private void ApplyDefaultCreationTheme() {
        Styles styles = (Styles)_wordprocessingDocument.MainDocumentPart!.StyleDefinitionsPart!.Styles!.CloneNode(true);
        ReplaceCreationTableStyle(styles, OfficeTheme.Default);
        ApplyThemeDefaults(styles, OfficeTheme.Default, paragraphSpacing: null);
    }

    private void ApplyThemeDefaults(Styles styles, OfficeTheme theme, SpacingBetweenLines? paragraphSpacing) {
        _creationTheme = theme;
        MainDocumentPart main = _wordprocessingDocument.MainDocumentPart!;
        StyleDefinitionsPart part = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        DocDefaults defaults = styles.GetFirstChild<DocDefaults>() ?? styles.PrependChild(new DocDefaults());
        RunPropertiesDefault runDefaults = defaults.RunPropertiesDefault ??= new RunPropertiesDefault();
        RunPropertiesBaseStyle run = runDefaults.RunPropertiesBaseStyle ??= new RunPropertiesBaseStyle();
        // Explicit projections retain native font-role ownership, including text added before the next theme switch.
        RunFonts fonts = paragraphSpacing != null ? WordThemeDefaults.BodyFonts() : new RunFonts();
        // Matching literal fallbacks also serve the existing configured-default getters and consumers.
        fonts.Ascii = theme.Fonts.Body.Family;
        fonts.HighAnsi = theme.Fonts.Body.Family;
        fonts.EastAsia = theme.Fonts.Body.Family;
        fonts.ComplexScript = theme.Fonts.Body.Family;
        run.RunFonts = fonts;
        string size = Math.Max(2, (int)Math.Round(theme.TypeScale.Body.ToPoints() * 2D, MidpointRounding.AwayFromZero))
            .ToString(System.Globalization.CultureInfo.InvariantCulture);
        run.FontSize = new FontSize { Val = size };
        run.FontSizeComplexScript = new FontSizeComplexScript { Val = size };
        run.Color = new Color { Val = WordThemeDefaults.PageText(theme).ToRgbHex() };
        if (paragraphSpacing != null) {
            ParagraphPropertiesDefault paragraphDefaults = defaults.ParagraphPropertiesDefault ??= new ParagraphPropertiesDefault();
            ParagraphPropertiesBaseStyle paragraph = paragraphDefaults.ParagraphPropertiesBaseStyle ??= new ParagraphPropertiesBaseStyle();
            WordThemeDefaults.SetVisualChild(paragraph, paragraphSpacing);
        }
        part.Styles = styles;
        ThemePart nativeTheme = main.ThemePart ?? main.AddNewPart<ThemePart>();
        if (nativeTheme.Theme == null) GenerateThemePart1Content(nativeTheme);
        ApplyCreationThemePart(nativeTheme, theme);
        if (paragraphSpacing != null && main.Document?.DocumentBackground is { } background &&
            (background.Color != null || background.ThemeColor != null || background.ThemeTint != null || background.ThemeShade != null)) {
            // Explicit typography targets Word's light page. Retain any background image children
            // and relationships, but remove theme selectors that would override the white RGB.
            background.Color = "FFFFFF";
            background.ThemeColor = null;
            background.ThemeTint = null;
            background.ThemeShade = null;
        }
    }
}
