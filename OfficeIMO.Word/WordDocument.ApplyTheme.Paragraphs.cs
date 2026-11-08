using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    private static void ApplyThemeParagraph(Paragraph paragraph, ThemeTextRole role, Style appearance, OfficeTheme theme) {
        ParagraphProperties properties = paragraph.ParagraphProperties ??= new ParagraphProperties();
        WordThemeDefaults.SetVisualChild(properties, appearance.StyleParagraphProperties!.SpacingBetweenLines!);
        if (role == ThemeTextRole.Code) {
            properties.Shading = new Shading { Val = ShadingPatternValues.Clear, Color = "auto",
                Fill = (theme.IsDark ? OfficeTheme.Modern.Surfaces.Surface : theme.Surfaces.Surface).Resolve(theme).ToRgbHex() };
        }
        if (role == ThemeTextRole.Title || role == ThemeTextRole.Footer) {
            OfficeThemeRule rule = role == ThemeTextRole.Title ? theme.TextStyle.TitleRule : theme.Footer.Rule;
            ParagraphBorders borders = properties.ParagraphBorders ??= new ParagraphBorders();
            if (role == ThemeTextRole.Title) {
                borders.BottomBorder = new BottomBorder { Val = rule.IsVisible ? BorderValues.Single : BorderValues.Nil,
                    Color = rule.Color.Resolve(theme).ToRgbHex(), Size = WordThemeDefaults.EighthPoints(rule.Width.ToPoints()) };
            } else {
                borders.TopBorder = new TopBorder { Val = rule.IsVisible ? BorderValues.Single : BorderValues.Nil,
                    Color = rule.Color.Resolve(theme).ToRgbHex(), Size = WordThemeDefaults.EighthPoints(rule.Width.ToPoints()) };
            }
        }
    }
}
