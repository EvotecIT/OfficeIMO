using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    // Formatting inspection is independent of any page renderer.
    internal WordCapsStyle ResolveEffectiveCapsStyle() => ResolveEffectiveCapsStyle(GetDirectCapsOverrides());

    internal WordCapsStyle ResolveEffectiveCapsStyle((bool? Caps, bool? SmallCaps) direct,
        WordStyleDefinitionsReader.StyleDefinitionSet? streamedStyles = null) {
        (bool? caps, bool? smallCaps) = direct;
        if ((!caps.HasValue || !smallCaps.HasValue) && streamedStyles != null) {
            ResolveStreamedCaps(streamedStyles, ref caps, ref smallCaps);
        } else if (!caps.HasValue || !smallCaps.HasValue) {
            foreach (StyleRunProperties properties in WordRunFormatting.EnumerateRunStyleProperties(this)) {
                caps ??= ReadCapsOnOff(properties.GetFirstChild<Caps>());
                smallCaps ??= ReadCapsOnOff(properties.GetFirstChild<SmallCaps>());
                if (caps.HasValue && smallCaps.HasValue) break;
            }

            if (!caps.HasValue || !smallCaps.HasValue) {
                RunPropertiesBaseStyle? defaults = _document?._wordprocessingDocument.MainDocumentPart?
                    .StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
                caps ??= ReadCapsOnOff(defaults?.GetFirstChild<Caps>());
                smallCaps ??= ReadCapsOnOff(defaults?.GetFirstChild<SmallCaps>());
            }
        }
        return caps == true ? WordCapsStyle.Caps : smallCaps == true ? WordCapsStyle.SmallCaps : WordCapsStyle.None;
    }

    /// <summary>
    /// The same cascade as the DOM path (character style chain, paragraph style chain, document defaults) over
    /// definitions streamed from an unloaded styles part.
    /// </summary>
    private void ResolveStreamedCaps(WordStyleDefinitionsReader.StyleDefinitionSet styles, ref bool? caps, ref bool? smallCaps) {
        if (!ApplyStreamedChain(styles, CharacterStyleId, "character", ref caps, ref smallCaps)) {
            ApplyStreamedChain(styles, StyleId, "paragraph", ref caps, ref smallCaps);
        }
        caps ??= styles.DefaultCaps;
        smallCaps ??= styles.DefaultSmallCaps;
    }

    // Mirrors WordRunFormatting.EnumerateStyleChain; returns true once both toggles are known.
    private static bool ApplyStreamedChain(WordStyleDefinitionsReader.StyleDefinitionSet styles, string? styleId, string type,
        ref bool? caps, ref bool? smallCaps) {
        IReadOnlyList<WordStyleDefinitionsReader.StyleDefinition> definitions = styles.Styles;
        if (string.IsNullOrWhiteSpace(styleId)) {
            styleId = null;
            foreach (WordStyleDefinitionsReader.StyleDefinition definition in definitions) {
                if (definition.IsDefault && string.Equals(definition.Type, type, StringComparison.Ordinal)) {
                    styleId = definition.StyleId;
                    break;
                }
            }
        }
        var visited = new HashSet<string>(StringComparer.Ordinal);
        while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
            WordStyleDefinitionsReader.StyleDefinition? style = null;
            foreach (WordStyleDefinitionsReader.StyleDefinition definition in definitions) {
                if (string.Equals(definition.StyleId, styleId, StringComparison.Ordinal) &&
                    (definition.Type == null || string.Equals(definition.Type, type, StringComparison.Ordinal))) {
                    style = definition;
                    break;
                }
            }
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
