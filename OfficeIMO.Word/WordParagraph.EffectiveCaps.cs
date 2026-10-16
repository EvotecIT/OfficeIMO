using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordParagraph {
    // Formatting inspection is independent of any page renderer.
    internal WordCapsStyle ResolveEffectiveCapsStyle() {
        RunProperties? direct = IsHyperLink ? Hyperlink?._runProperties : _runProperties;
        bool caps = ResolveEffectiveOnOff<Caps>(direct?.Caps);
        bool smallCaps = ResolveEffectiveOnOff<SmallCaps>(direct?.SmallCaps);
        return caps ? WordCapsStyle.Caps : smallCaps ? WordCapsStyle.SmallCaps : WordCapsStyle.None;
    }

    private bool ResolveEffectiveOnOff<T>(T? direct) where T : OnOffType {
        if (direct != null) return direct.Val?.Value ?? true;
        Styles? styles = _document?._wordprocessingDocument.MainDocumentPart?.StyleDefinitionsPart?.Styles;
        foreach (StyleRunProperties properties in WordRunFormatting.EnumerateRunStyleProperties(this)) {
            T? value = properties.GetFirstChild<T>();
            if (value != null) return value.Val?.Value ?? true;
        }
        T? fallback = styles?.DocDefaults?.RunPropertiesDefault?.RunPropertiesBaseStyle?.GetFirstChild<T>();
        return fallback != null && (fallback.Val?.Value ?? true);
    }

}
