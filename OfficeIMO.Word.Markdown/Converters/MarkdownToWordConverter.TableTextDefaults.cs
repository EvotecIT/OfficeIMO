using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown;

internal partial class MarkdownToWordConverter {
    /// <summary>Seeds a fresh cell paragraph from the native role owner before authored inline formatting.</summary>
    private static StyleRunProperties? PrepareNativeTableTextDefaults(WordParagraph paragraph, Omd.MarkdownVisualTheme? theme) {
        if (paragraph._paragraph.Ancestors<TableCell>().FirstOrDefault() is not { } cell ||
            theme?.EffectiveSourceTheme is not { } source) return null;
        // Explicit block roles retain their own native typography and spacing.
        string? style = paragraph._paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value;
        if (!string.IsNullOrEmpty(style) && style is not ("Normal" or "ListParagraph")) return null;
        Style appearance = WordThemeDefaults.CreateTableTextStyle(source, WordDocument.IsThemeHeaderCell(cell));
        ParagraphProperties properties = paragraph._paragraph.ParagraphProperties ??= new ParagraphProperties();
        properties.SpacingBetweenLines ??= (SpacingBetweenLines)appearance.StyleParagraphProperties!.SpacingBetweenLines!.CloneNode(true);
        return appearance.StyleRunProperties;
    }

    private static InlineFormatState ResolveNativeTextContext(WordParagraph paragraph, MarkdownWordExportOptions options, InlineFormatState? format = null) =>
        format ?? CreateDefaultInlineFormatState(PrepareNativeTableTextDefaults(paragraph, ResolveTheme(options)));

    /// <summary>Emits fallback or literal text through the same native cell context as parsed inlines.</summary>
    private static WordParagraph AddContextText(WordParagraph paragraph, string? text, MarkdownWordExportOptions options, InlineFormatState? format = null) {
        WordParagraph run = paragraph.AddText(text ?? string.Empty);
        ApplyNativeContextRunFormatting(run, paragraph, options, format);
        return run;
    }

    private static void ApplyNativeContextRunFormatting(WordParagraph run, WordParagraph paragraph, MarkdownWordExportOptions options, InlineFormatState? format = null) {
        InlineFormatState context = ResolveNativeTextContext(paragraph, options, format);
        if (context.NativeDefaults != null) ApplyInlineRunFormatting(run, context, ResolveDefaultFontFamily(options), defaultTextColorHex: null);
    }

    /// <summary>Projects fallback native properties only onto the newly emitted run; authored overrides follow.</summary>
    private static void ApplyNativeTableRunDefaults(WordParagraph run, InlineFormatState format, bool omitFont) {
        if (format.NativeDefaults == null) return;
        IEnumerable<Run> nativeRuns = run._run is { } text ? new[] { text }
            : run._hyperlink?.Elements<Run>() ?? Enumerable.Empty<Run>();
        foreach (Run native in nativeRuns) {
            RunProperties properties = native.RunProperties ??= new RunProperties();
            foreach (OpenXmlElement property in format.NativeDefaults.ChildElements) {
                // Normal supplies no bold default; the native table role can keep owning its toggle.
                if (property is Bold or BoldComplexScript) continue;
                if (omitFont && property is RunFonts) continue;
                WordThemeDefaults.SetVisualChild(properties, property);
            }
        }
    }
}
