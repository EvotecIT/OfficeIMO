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

    /// <summary>Applies only creation-owned table/default-font tokens and native theme content to this package.</summary>
    private void ApplyCreationTheme(OfficeTheme theme) {
        _creationTheme = theme;
        MainDocumentPart main = _wordprocessingDocument.MainDocumentPart!;
        StyleDefinitionsPart part = main.StyleDefinitionsPart ?? main.AddNewPart<StyleDefinitionsPart>();
        // Clone the package's catalog, never the static shared catalog's caller-specific contents.
        Styles styles = part.Styles != null ? (Styles)part.Styles.CloneNode(true) : (Styles)BuiltInStyleDefinitions.Value.CloneNode(true);
        ReplaceCreationTableStyle(styles, theme);
        DocDefaults defaults = styles.GetFirstChild<DocDefaults>() ?? styles.PrependChild(new DocDefaults());
        RunPropertiesDefault runDefaults = defaults.RunPropertiesDefault ??= new RunPropertiesDefault();
        RunPropertiesBaseStyle run = runDefaults.RunPropertiesBaseStyle ??= new RunPropertiesBaseStyle();
        run.RunFonts = new RunFonts { Ascii = theme.Fonts.Body.Family, HighAnsi = theme.Fonts.Body.Family,
            EastAsia = theme.Fonts.Body.Family, ComplexScript = theme.Fonts.Body.Family };
        string size = Math.Max(2, (int)Math.Round(theme.TypeScale.Body.ToPoints() * 2D, MidpointRounding.AwayFromZero))
            .ToString(System.Globalization.CultureInfo.InvariantCulture);
        run.FontSize = new FontSize { Val = size };
        run.FontSizeComplexScript = new FontSizeComplexScript { Val = size };
        run.Color = new Color { Val = WordThemeDefaults.PageText(theme).ToRgbHex() };
        part.Styles = styles;
        ThemePart nativeTheme = main.ThemePart ?? main.AddNewPart<ThemePart>();
        if (nativeTheme.Theme == null) GenerateThemePart1Content(nativeTheme);
        ApplyCreationThemePart(nativeTheme, theme);
    }
}
