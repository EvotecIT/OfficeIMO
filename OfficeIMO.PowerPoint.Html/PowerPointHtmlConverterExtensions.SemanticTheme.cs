using OfficeIMO.Html;
using PptCore = OfficeIMO.PowerPoint;

namespace OfficeIMO.PowerPoint.Html;

public static partial class PowerPointHtmlConverterExtensions {
    // Source run and table colors stay native, so their panel needs the source palette too.
    private static void AppendSemanticSlideStyle(StringBuilder body, PptCore.PowerPointSlide slide) {
        OfficeTheme theme = slide.AuthoringTheme;
        var background = slide.GetBackgroundColors();
        if (background.Kind == PptCore.PowerPointSlideBackgroundKind.SolidColor && OfficeColor.TryParse(background.Color, out var color))
            theme = theme.WithPalette(theme.Palette.With(background: color));
        body.Append(" style=\"")
            .Append(OfficeHtmlText.EscapeAttribute(OfficeHtmlDocumentShell.GetThemeProperties(theme)))
            .Append("background-color:var(--officeimo-page-background);color:var(--officeimo-text);\"");
    }
}
