using OfficeIMO.OpenXml.Internal;
using Omd = OfficeIMO.Markdown;

namespace OfficeIMO.Word.Markdown;

internal partial class MarkdownToWordConverter {
    /// <summary>Seeds a fresh cell paragraph from the native role owner before authored inline formatting.</summary>
    private static OfficeXmlNode PrepareNativeTableTextDefaults(WordParagraph paragraph, Omd.MarkdownVisualTheme? theme) {
        OfficeXmlNode element = paragraph.Node();
        OfficeXmlNode cell = element.Ancestor(WordXml.W, "tc");
        if (cell.IsNone || theme?.EffectiveSourceTheme is not { } source) return default;
        // Explicit block roles retain their own native typography and spacing.
        string? style = element.WChild("pPr").ChildVal("pStyle");
        if (!string.IsNullOrEmpty(style) && paragraph.BuiltInStyleId is not ("Normal" or "ListParagraph")) return default;
        OfficeXmlNode appearance = WordConverterNodes.CreateTableTextStyleNode(source, WordConverterNodes.IsThemeHeaderCell(cell));
        OfficeXmlNode properties = element.GetOrAddChild(WordXml.W, "pPr", WPView.Order);
        if (!properties.WChild("spacing").Exists) {
            properties.InsertOrdered(appearance.WChild("pPr").WChild("spacing").Clone(), WPPrView.Order);
        }
        return appearance.WChild("rPr");
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
        if (context.NativeDefaults.Exists) ApplyInlineRunFormatting(run, context, ResolveDefaultFontFamily(options), defaultTextColorHex: null);
    }

    /// <summary>Projects fallback native properties only onto the newly emitted run; authored overrides follow.</summary>
    private static void ApplyNativeTableRunDefaults(WordParagraph run, InlineFormatState format, bool omitFont) {
        if (format.NativeDefaults.IsNone) return;
        OfficeXmlNode text = run.RunNode();
        OfficeXmlNode hyperlink = run.HyperlinkNode();
        IEnumerable<OfficeXmlNode> nativeRuns = text.Exists ? new[] { text }
            : hyperlink.Exists ? hyperlink.WElements("r") : Enumerable.Empty<OfficeXmlNode>();
        foreach (OfficeXmlNode native in nativeRuns) {
            OfficeXmlNode properties = native.GetOrAddChild(WordXml.W, "rPr", WRView.Order);
            foreach (OfficeXmlNode property in format.NativeDefaults.Elements()) {
                // Normal supplies no bold default; the native table role can keep owning its toggle.
                if (property.IsW("b") || property.IsW("bCs")) continue;
                if (omitFont && property.IsW("rFonts")) continue;
                WordConverterNodes.SetVisualChild(properties, property);
            }
        }
    }
}
