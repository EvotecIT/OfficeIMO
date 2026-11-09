using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    // Formatting inspection is independent of any page renderer.
    internal WordCapsStyle ResolveEffectiveCapsStyle() => ResolveEffectiveCapsStyle(GetDirectCapsOverrides());

    internal WordCapsStyle ResolveEffectiveCapsStyle((bool? Caps, bool? SmallCaps) direct) {
        (bool? caps, bool? smallCaps) = direct;
        if (!caps.HasValue || !smallCaps.HasValue) {
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

    /// <summary>Reads direct caps toggles without applying inheritance; absent toggles remain null.</summary>
    internal (bool? Caps, bool? SmallCaps) GetDirectCapsOverrides() {
        RunProperties? direct = IsHyperLink ? Hyperlink?._runProperties : _runProperties;
        return (ReadCapsOnOff(direct?.Caps), ReadCapsOnOff(direct?.SmallCaps));
    }

    private static bool? ReadCapsOnOff(OnOffType? value) => value == null ? null : value.Val?.Value ?? true;
}
