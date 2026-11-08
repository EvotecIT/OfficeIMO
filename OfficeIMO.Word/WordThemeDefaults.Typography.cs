using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

internal static partial class WordThemeDefaults {
    internal static RunFonts BodyFonts() => new() {
        AsciiTheme = ThemeFontValues.MinorAscii, HighAnsiTheme = ThemeFontValues.MinorHighAnsi,
        EastAsiaTheme = ThemeFontValues.MinorEastAsia, ComplexScriptTheme = ThemeFontValues.MinorBidi
    };

    internal static Style CreateBody(OfficeTheme theme) => new(
        new StyleParagraphProperties(new SpacingBetweenLines {
            Before = Twips(theme.Spacing.ParagraphBefore.ToPoints()), After = Twips(theme.Spacing.ParagraphAfter.ToPoints()),
            Line = Math.Round(theme.TextStyle.LineSpacing * 240D, MidpointRounding.AwayFromZero).ToString(System.Globalization.CultureInfo.InvariantCulture),
            LineRule = LineSpacingRuleValues.Auto
        }),
        new StyleRunProperties(BodyFonts(), new Color { Val = PageText(theme).ToRgbHex() },
            new FontSize { Val = HalfPoints(theme.TypeScale.Body) },
            new FontSizeComplexScript { Val = HalfPoints(theme.TypeScale.Body) }));

    /// <summary>Creates table body/header text defaults for both new content and explicit document restyling.</summary>
    internal static Style CreateTableTextStyle(OfficeTheme theme, bool header) {
        Style appearance = CreateTableStyle(theme);
        appearance.StyleRunProperties!.RunFonts = BodyFonts();
        if (header) {
            RunPropertiesBaseStyle properties = appearance.Elements<TableStyleProperties>()
                .Single(part => part.Type?.Value == TableStyleOverrideValues.FirstRow).RunPropertiesBaseStyle!;
            foreach (OpenXmlElement property in properties.ChildElements) SetVisualChild(appearance.StyleRunProperties, property);
        }
        return appearance;
    }

    /// <summary>Updates existing role definitions without changing style identity, inheritance, or default flags.</summary>
    internal static void ApplyCreationTypography(Styles styles, OfficeTheme theme) {
        foreach (Style existing in styles.Elements<Style>()) {
            string? id = existing.StyleId?.Value;
            Style? appearance = null;
            if (id == "Normal" || existing.Type?.Value == StyleValues.Paragraph && existing.Default?.Value == true)
                appearance = CreateBody(theme);
            else if (id == "Title") appearance = CreateTitle(theme);
            else if (id == "Subtitle") appearance = CreateSubtitle(theme);
            else if (TryHeadingLevel(id, out int level)) appearance = CreateHeading(level, theme);
            else if (id != null && id.EndsWith("Char", StringComparison.Ordinal) && TryHeadingLevel(id.Substring(0, id.Length - 4), out level))
                appearance = CreateHeadingCharacter(level, theme);
            if (appearance == null) continue;

            StyleRunProperties run = existing.StyleRunProperties ??= new StyleRunProperties();
            foreach (OpenXmlElement child in appearance.StyleRunProperties!.ChildElements) SetVisualChild(run, child);
            if (appearance.StyleParagraphProperties?.SpacingBetweenLines is { } spacing) {
                StyleParagraphProperties paragraph = existing.StyleParagraphProperties ??= new StyleParagraphProperties();
                SetVisualChild(paragraph, spacing);
            }
        }
    }

    internal static bool TryHeadingLevel(string? id, out int level) {
        level = 0;
        if (id is not { Length: 8 } || !id.StartsWith("Heading", StringComparison.Ordinal) || id[7] < '1' || id[7] > '9') return false;
        level = id[7] - '0';
        return true;
    }

    internal static void SetVisualChild(OpenXmlCompositeElement target, OpenXmlElement value) {
        foreach (OpenXmlElement existing in target.ChildElements.Where(child => child.LocalName == value.LocalName && child.NamespaceUri == value.NamespaceUri).ToArray())
            existing.Remove();
        target.AddChild(value.CloneNode(true), true);
    }
}
