using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    /// <summary>Adds a shaded callout title and body, returning its title paragraph.</summary>
    public WordParagraph AddCallout(OfficeCalloutKind kind, string title, string body, OfficeTheme? theme = null) {
        ArgumentNullException.ThrowIfNull(title);
        ArgumentNullException.ThrowIfNull(body);
        OfficeColor color = kind switch {
            OfficeCalloutKind.Info => OfficeColor.Accent1,
            OfficeCalloutKind.Success => OfficeColor.Success,
            OfficeCalloutKind.Warning => OfficeColor.Warning,
            OfficeCalloutKind.Danger => OfficeColor.Danger,
            _ => throw new ArgumentOutOfRangeException(nameof(kind))
        };
        color = color.Resolve(theme ?? OfficeTheme.Default);
        var heading = AddParagraph();
        heading.ShadingFillColor = color;
        heading.AddFormattedText(title, bold: true);
        var paragraph = AddParagraph(body);
        paragraph.ShadingFillColor = color;
        return heading;
    }

    /// <summary>Adds a code paragraph using a named monospace paragraph style. The language is a descriptive hint.</summary>
    public WordParagraph AddCodeBlock(string code, string? language = null) {
        ArgumentNullException.ThrowIfNull(code);
        const string styleId = "OfficeIMOCodeBlock";
        var font = WordFontResolver.Resolve("monospace") ?? "Consolas";
        if (!Styles.Contains(styleId))
            Styles.AddParagraphStyle(styleId, Styles.FindByName("Code Block") == null ? "Code Block" : null).FontFamily = font;
        return AddParagraph(code).SetStyleId(styleId);
    }
}
