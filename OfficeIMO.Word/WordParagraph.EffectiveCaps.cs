using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    // Formatting inspection is independent of any page renderer.
    internal WordCapsStyle ResolveEffectiveCapsStyle() {
        RunProperties? direct = IsHyperLink ? Hyperlink?._runProperties : _runProperties;
        bool? caps = ReadOnOff(direct?.Caps);
        bool? smallCaps = ReadOnOff(direct?.SmallCaps);
        if (!caps.HasValue || !smallCaps.HasValue) {
            foreach (StyleRunProperties properties in WordRunFormatting.EnumerateRunStyleProperties(this)) {
                caps ??= ReadOnOff(properties.GetFirstChild<Caps>());
                smallCaps ??= ReadOnOff(properties.GetFirstChild<SmallCaps>());
                if (caps.HasValue && smallCaps.HasValue) break;
            }

            if (!caps.HasValue || !smallCaps.HasValue) {
                RunPropertiesBaseStyle? defaults = _document?._wordprocessingDocument.MainDocumentPart?
                    .StyleDefinitionsPart?.Styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle;
                caps ??= ReadOnOff(defaults?.GetFirstChild<Caps>());
                smallCaps ??= ReadOnOff(defaults?.GetFirstChild<SmallCaps>());
            }
        }
        return caps == true ? WordCapsStyle.Caps : smallCaps == true ? WordCapsStyle.SmallCaps : WordCapsStyle.None;

        static bool? ReadOnOff(OnOffType? value) => value == null ? null : value.Val?.Value ?? true;
    }

}
