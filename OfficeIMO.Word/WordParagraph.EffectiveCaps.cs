using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    // Formatting inspection is independent of any page renderer.
    internal WordCapsStyle ResolveEffectiveCapsStyle() => ResolveEffectiveCapsStyle(GetDirectCapsOverrides());

    internal WordCapsStyle ResolveEffectiveCapsStyle((bool? Caps, bool? SmallCaps) direct) {
        (bool? caps, bool? smallCaps) = direct;
        if ((!caps.HasValue || !smallCaps.HasValue) && _document != null) {
            WordStyleCatalog styles = WordStyleCatalog.For(_document);
            if (styles.IsStreamed) {
                // The streamed facts answer caps without parsing any style's properties.
                ResolveStreamedCaps(styles, ref caps, ref smallCaps);
            } else {
                foreach (StyleRunProperties properties in WordRunFormatting.EnumerateRunStyleProperties(this)) {
                    caps ??= ReadCapsOnOff(properties.GetFirstChild<Caps>());
                    smallCaps ??= ReadCapsOnOff(properties.GetFirstChild<SmallCaps>());
                    if (caps.HasValue && smallCaps.HasValue) break;
                }

                if (!caps.HasValue || !smallCaps.HasValue) {
                    RunPropertiesBaseStyle? defaults = styles.DocumentDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
                    caps ??= ReadCapsOnOff(defaults?.GetFirstChild<Caps>());
                    smallCaps ??= ReadCapsOnOff(defaults?.GetFirstChild<SmallCaps>());
                }
            }
        }
        return caps == true ? WordCapsStyle.Caps : smallCaps == true ? WordCapsStyle.SmallCaps : WordCapsStyle.None;
    }

    /// <summary>
    /// The same cascade as the DOM path (character style chain, paragraph style chain, document defaults) over
    /// definitions streamed from an unloaded styles part.
    /// </summary>
    private void ResolveStreamedCaps(WordStyleCatalog styles, ref bool? caps, ref bool? smallCaps) {
        if (!ApplyStreamedChain(styles, CharacterStyleId, StyleValues.Character, ref caps, ref smallCaps)) {
            ApplyStreamedChain(styles, StyleId, StyleValues.Paragraph, ref caps, ref smallCaps);
        }
        caps ??= styles.Streamed!.DefaultCaps;
        smallCaps ??= styles.Streamed!.DefaultSmallCaps;
    }

    // Mirrors WordStyleCatalog.EnumerateChain over the streamed facts; returns true once both toggles are known.
    private static bool ApplyStreamedChain(WordStyleCatalog styles, string? styleId, StyleValues type,
        ref bool? caps, ref bool? smallCaps) {
        if (string.IsNullOrWhiteSpace(styleId)) styleId = styles.DefaultStyleId(type);
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
            WordStyleDefinitionsReader.StyleDefinition? style = styles.FindDefinition(styleId, type);
            if (style == null) return false;
            if (style.HasRunProperties) {
                caps ??= style.Caps;
                smallCaps ??= style.SmallCaps;
                if (caps.HasValue && smallCaps.HasValue) return true;
            }
            styleId = style.BasedOn;
        }
        return false;
    }

    /// <summary>Reads direct caps toggles without applying inheritance; absent toggles remain null.</summary>
    internal (bool? Caps, bool? SmallCaps) GetDirectCapsOverrides() {
        RunProperties? direct = IsHyperLink ? Hyperlink?._runProperties : _runProperties;
        return (ReadCapsOnOff(direct?.Caps), ReadCapsOnOff(direct?.SmallCaps));
    }

    private static bool? ReadCapsOnOff(OnOffType? value) => value == null ? null : value.Val?.Value ?? true;
}
