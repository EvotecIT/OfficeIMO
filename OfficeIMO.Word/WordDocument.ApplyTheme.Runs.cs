using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;
using OfficeIMO.Drawing;
using A = DocumentFormat.OpenXml.Drawing;

namespace OfficeIMO.Word;

public partial class WordDocument {
    // Native theme font references retain font ownership. Bold has no equivalent reference, so remember only
    // the two role-level values this operation wrote; a subsequent authored change takes precedence.
    private sealed class AppliedThemeBold {
        internal bool? Latin;
        internal bool? Complex;
    }

    private void ApplyThemeRun(Paragraph paragraph, Run run, ThemeTextRole role, Style appearance, Styles? styles, OfficeTheme theme) {
        Style[] characterStyles = WordRunFormatting.EnumerateStyleChain(styles, run.RunProperties?.RunStyle?.Val?.Value, StyleValues.Character).ToArray();
        Style[] paragraphStyles = WordRunFormatting.EnumerateStyleChain(styles, paragraph.ParagraphProperties?.ParagraphStyleId?.Val?.Value, StyleValues.Paragraph).ToArray();
        OpenXmlCompositeElement[] tableSources = WordRunFormatting.EnumerateTableRunPropertySources(paragraph, styles).ToArray();
        RunPropertiesBaseStyle? defaults = styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
        var sources = new List<OpenXmlCompositeElement>();
        if (run.RunProperties != null) sources.Add(run.RunProperties);
        sources.AddRange(characterStyles.Concat(paragraphStyles).Select(style => style.StyleRunProperties).OfType<OpenXmlCompositeElement>());
        sources.AddRange(tableSources);
        if (defaults != null) sources.Add(defaults);
        RunFonts fonts = ResolveThemeSourceFonts(sources);
        Text[] text = run.Elements<Text>().ToArray();
        bool plainAscii = text.Length > 0 && text.All(value => value.Text.All(character => character <= 0x7F));
        bool scriptOverride = fonts.Hint?.Value is { } hint && hint != FontTypeHintValues.Default ||
            Enabled(paragraph.ParagraphProperties?.BiDi) || paragraphStyles.Any(style => Enabled(style.StyleParagraphProperties?.GetFirstChild<BiDi>())) ||
            Enabled(styles?.DocDefaults?.ParagraphPropertiesDefault?.ParagraphPropertiesBaseStyle?.GetFirstChild<BiDi>()) ||
            sources.Any(source => Enabled(source.GetFirstChild<RightToLeftText>()) || Enabled(source.GetFirstChild<ComplexScript>()));
        bool asciiSpecified = !string.IsNullOrWhiteSpace(fonts.Ascii?.Value) ||
            !string.IsNullOrEmpty(ReadThemeFontAttribute(fonts, "asciiTheme").Value);
        // An unused hAnsi/Asian/complex-script face must not turn ordinary ASCII text into code.
        // Keep the conservative slot set when text or explicit script settings need those slots.
        IEnumerable<string> fontSlots = plainAscii && !scriptOverride && asciiSpecified ? ThemeFontSlots.Take(1) : ThemeFontSlots;
        bool preserveFont = role == ThemeTextRole.Code || characterStyles.Any(style => FormattingHelper.IsCodeCharacterStyle(style.StyleId?.Value)) ||
            run.Ancestors<SdtRun>().Any(control => control.SdtProperties?.GetFirstChild<DocumentFormat.OpenXml.Office2010.Word.SdtContentCheckBox>() != null) ||
            HasAuthoredSemanticFont(fonts, fontSlots);
        RunProperties properties = run.RunProperties ??= new RunProperties();
        if (!run.Elements<SymbolChar>().Any()) {
            RunFonts targetFonts = preserveFont ? FreezeThemeFontReferences(fonts) :
                (RunFonts)(appearance.StyleRunProperties?.RunFonts is { } roleFonts && role != ThemeTextRole.Table && role != ThemeTextRole.TableHeader
                    ? roleFonts.CloneNode(true) : WordThemeDefaults.BodyFonts());
            targetFonts.Hint = fonts.Hint;
            WordThemeDefaults.SetVisualChild(properties, targetFonts);
        }
        StyleRunProperties visual = appearance.StyleRunProperties!;
        WordThemeDefaults.SetVisualChild(properties, visual.FontSize!);
        WordThemeDefaults.SetVisualChild(properties, visual.FontSizeComplexScript!);
        bool link = run.Ancestors<Hyperlink>().Any() || characterStyles.Any(style => style.StyleId?.Value is "Hyperlink" or "FollowedHyperlink");
        WordThemeDefaults.SetVisualChild(properties, link ? new Color { Val = theme.TextStyle.LinkColor.Resolve(theme).ToRgbHex() } : visual.Color!);

        AppliedThemeBold ownership = run.Annotation<AppliedThemeBold>() ?? new AppliedThemeBold();
        ownership.Latin = ApplyThemeRoleBold(properties, visual.Bold, ownership.Latin, characterStyles, paragraphStyles, tableSources, defaults, role);
        ownership.Complex = ApplyThemeRoleBold(properties, visual.BoldComplexScript, ownership.Complex, characterStyles, paragraphStyles, tableSources, defaults, role);
        run.RemoveAnnotations<AppliedThemeBold>();
        run.AddAnnotation(ownership);

        static bool Enabled(OnOffType? value) => value != null && (value.Val?.Value ?? true);
    }

    private static bool? ApplyThemeRoleBold<T>(RunProperties properties, T? appearance, bool? previous,
        Style[] characterStyles, Style[] paragraphStyles, OpenXmlCompositeElement[] tableSources,
        RunPropertiesBaseStyle? defaults, ThemeTextRole role) where T : OnOffType, new() {
        T? direct = properties.GetFirstChild<T>();
        bool owned = previous.HasValue && direct != null && (direct.Val?.Value ?? true) == previous.Value;
        if (direct != null && !owned) return null;
        if (owned) direct!.Remove();
        bool heading = role >= ThemeTextRole.Heading1 && role <= ThemeTextRole.Heading9 || role == ThemeTextRole.Title;
        Style[] authoredCharacters = characterStyles.Where(style => !heading || !IsThemeHeadingCharacter(style.StyleId?.Value)).ToArray();
        IEnumerable<Style> authoredParagraphs = paragraphStyles.Where(style => !heading || style.StyleId?.Value != "Title" && !WordThemeDefaults.TryHeadingLevel(style.StyleId?.Value, out _));
        bool tableRole = role == ThemeTextRole.Table || role == ThemeTextRole.TableHeader;
        bool hasAuthoredReset = authoredParagraphs.Any(style => Read(style.StyleRunProperties).HasValue) ||
            tableSources.Any(source => (!tableRole || source.Ancestors<Style>().FirstOrDefault()?.StyleId?.Value != WordThemeDefaults.TableStyleId) && Read(source).HasValue);
        bool? character = ResolveCharacterToggle(characterStyles.Select(style => style.StyleRunProperties));
        bool enabled = Read(defaults) == true;
        if (enabled || hasAuthoredReset || ResolveCharacterToggle(authoredCharacters.Select(style => style.StyleRunProperties)) == true) {
            // Word resets paragraph/table declarations instead of toggling the previous
            // value (MS-OI29500 2.1.258; MS-OE376 2.1.248). The nearest declaration
            // therefore wins within those chains. Character styles retain their distinct
            // toggle hierarchy; direct run properties remain absolute.
            bool? table = tableSources.Select(Read).FirstOrDefault(value => value.HasValue);
            bool? paragraph = paragraphStyles.Select(style => Read(style.StyleRunProperties)).FirstOrDefault(value => value.HasValue);
            if (table.HasValue) enabled = table.Value;
            if (paragraph.HasValue) enabled = paragraph.Value;
            if (character == true) enabled = !enabled;
            WordThemeDefaults.SetVisualChild(properties, new T { Val = enabled });
            return null;
        }
        if (appearance == null) return null;
        WordThemeDefaults.SetVisualChild(properties, appearance);
        return appearance.Val?.Value ?? true;

        static bool? Read(OpenXmlCompositeElement? source) =>
            source?.GetFirstChild<T>() is { } value ? value.Val?.Value ?? true : null;

        static bool? ResolveCharacterToggle(IEnumerable<OpenXmlCompositeElement?> sources) {
            bool? result = null;
            foreach (OpenXmlCompositeElement? source in sources) {
                if (Read(source) is not bool declared) continue;
                result ??= false;
                if (declared) result = !result.Value;
            }
            return result;
        }
    }

    private static bool IsThemeHeadingCharacter(string? id) => id != null && id.EndsWith("Char", StringComparison.Ordinal) &&
        WordThemeDefaults.TryHeadingLevel(id.Substring(0, id.Length - 4), out _);

    private static OpenXmlAttribute ReadThemeFontAttribute(RunFonts fonts, string name) =>
        fonts.GetAttributes().FirstOrDefault(attribute => attribute.LocalName == name && attribute.NamespaceUri == fonts.NamespaceUri);

    private static readonly string[] ThemeFontSlots = { "ascii", "hAnsi", "eastAsia", "cs" };

    private static RunFonts ResolveThemeSourceFonts(IEnumerable<OpenXmlCompositeElement> sources) {
        RunFonts[] fonts = sources.Select(source => source.GetFirstChild<RunFonts>()).OfType<RunFonts>().ToArray();
        var resolved = new RunFonts { Hint = fonts.Select(font => font.Hint).FirstOrDefault(value => value != null) };
        foreach (string slot in ThemeFontSlots) {
            string selector = slot == "cs" ? "cstheme" : slot + "Theme";
            foreach (RunFonts source in fonts) {
                OpenXmlAttribute literal = ReadThemeFontAttribute(source, slot);
                OpenXmlAttribute theme = ReadThemeFontAttribute(source, selector);
                if (string.IsNullOrEmpty(literal.Value) && string.IsNullOrEmpty(theme.Value)) continue;
                if (!string.IsNullOrEmpty(literal.Value)) resolved.SetAttribute(literal);
                if (!string.IsNullOrEmpty(theme.Value)) resolved.SetAttribute(theme);
                break;
            }
        }
        return resolved;
    }

    private bool HasAuthoredSemanticFont(RunFonts fonts, IEnumerable<string> slots) {
        foreach (string slot in slots) {
            string selector = slot == "cs" ? "cstheme" : slot + "Theme";
            if (!string.IsNullOrEmpty(ReadThemeFontAttribute(fonts, selector).Value)) continue;
            string? family = ReadThemeFontAttribute(fonts, slot).Value;
            if (string.IsNullOrWhiteSpace(family)) continue;
            if (OfficeDocumentFontSubstitutions.PreservesSemanticFont(family!)) return true;
            Font? declared = _wordprocessingDocument.MainDocumentPart?.FontTablePart?.Fonts?.Elements<Font>()
                .FirstOrDefault(font => string.Equals(font.Name?.Value, family, StringComparison.OrdinalIgnoreCase));
            if (declared?.GetFirstChild<Pitch>()?.Val?.Value == FontPitchValues.Fixed ||
                declared?.GetFirstChild<FontCharSet>()?.Val?.Value == "02") return true;
        }
        return false;
    }

    private RunFonts FreezeThemeFontReferences(RunFonts fonts) {
        RunFonts frozen = (RunFonts)fonts.CloneNode(true);
        A.FontScheme? scheme = _wordprocessingDocument.MainDocumentPart?.ThemePart?.Theme?.ThemeElements?.FontScheme;
        foreach (string slot in ThemeFontSlots) {
            string selector = slot == "cs" ? "cstheme" : slot + "Theme";
            string? value = ReadThemeFontAttribute(fonts, selector).Value;
            A.FontCollectionType? collection = value?.StartsWith("major", StringComparison.Ordinal) == true ? scheme?.MajorFont :
                value?.StartsWith("minor", StringComparison.Ordinal) == true ? scheme?.MinorFont : null;
            string? family = slot == "eastAsia" ? collection?.GetFirstChild<A.EastAsianFont>()?.Typeface?.Value :
                slot == "cs" ? collection?.GetFirstChild<A.ComplexScriptFont>()?.Typeface?.Value : collection?.GetFirstChild<A.LatinFont>()?.Typeface?.Value;
            if (string.IsNullOrWhiteSpace(family)) family = collection?.GetFirstChild<A.LatinFont>()?.Typeface?.Value;
            if (string.IsNullOrWhiteSpace(family)) continue;
            frozen.RemoveAttribute(selector, frozen.NamespaceUri);
            frozen.SetAttribute(new OpenXmlAttribute("w", slot, frozen.NamespaceUri, family));
        }
        string? fallback = frozen.Ascii?.Value ?? frozen.HighAnsi?.Value;
        if (!string.IsNullOrWhiteSpace(fallback)) {
            foreach (string slot in ThemeFontSlots) {
                string selector = slot == "cs" ? "cstheme" : slot + "Theme";
                if (string.IsNullOrEmpty(ReadThemeFontAttribute(frozen, slot).Value) && string.IsNullOrEmpty(ReadThemeFontAttribute(frozen, selector).Value))
                    frozen.SetAttribute(new OpenXmlAttribute("w", slot, frozen.NamespaceUri, fallback));
            }
        }
        return frozen;
    }
}
