namespace OfficeIMO.Rtf;

public sealed partial class RtfDocument {
    private sealed class MergeResourceMap {
        internal readonly Dictionary<int, int> ListDefinitions = new Dictionary<int, int>();
        internal readonly Dictionary<int, int> ListInstances = new Dictionary<int, int>();
        internal readonly Dictionary<RtfStyleKind, Dictionary<int, int>> Styles = new Dictionary<RtfStyleKind, Dictionary<int, int>>();
        internal readonly HashSet<object> Content = new HashSet<object>();
        internal int DefaultParagraphStyle;
        internal int? DefaultFont;
        internal int? DefaultLanguage;

        internal int? Style(int? id, RtfStyleKind kind) => id.HasValue && Styles.TryGetValue(kind, out Dictionary<int, int>? ids) && ids.TryGetValue(id.Value, out int value) ? value : null;
    }

    private MergeResourceMap ImportBindings(RtfDocument source, Dictionary<int, int> fonts, Dictionary<int, int> colors) {
        Writing.RtfDocumentWriter.EffectiveListTables lists = Writing.RtfDocumentWriter.BuildEffectiveListTables(source);
        source.ReplaceListDefinitions(lists.Definitions);
        source.ReplaceListOverrides(lists.Overrides);
        var map = new MergeResourceMap {
            DefaultFont = MapIndex(source.Settings.DefaultFontId ?? 0, fonts),
            DefaultLanguage = source.Settings.DefaultLanguageId
        };
        int definitionId = _listDefinitions.Count == 0 ? 1 : checked(_listDefinitions.Max(item => item.Id) + 1);
        foreach (RtfListDefinition definition in source.ListDefinitions) {
            map.ListDefinitions.Add(definition.Id, definitionId);
            definition.Id = definitionId++;
            _listDefinitions.Add(definition);
        }
        int instanceId = _listOverrides.Count == 0 ? 1 : checked(_listOverrides.Max(item => item.Id) + 1);
        foreach (RtfListOverride instance in source.ListOverrides) {
            map.ListInstances.Add(instance.Id, instanceId);
            instance.Id = instanceId++;
            instance.ListId = MapIndex(instance.ListId, map.ListDefinitions)!.Value;
            _listOverrides.Add(instance);
        }
        int styleId = _styles.Count == 0 ? 1 : checked(_styles.Max(item => item.Id) + 1);
        foreach (RtfStyle style in source.Styles) {
            if (!map.Styles.TryGetValue(style.Kind, out Dictionary<int, int>? ids)) {
                ids = new Dictionary<int, int>();
                map.Styles.Add(style.Kind, ids);
            }
            ids.Add(style.Id, styleId++);
        }
        foreach (RtfStyle style in source.Styles) {
            style.Id = map.Style(style.Id, style.Kind)!.Value;
            style.BasedOnStyleId = map.Style(style.BasedOnStyleId, style.Kind);
            style.NextStyleId = map.Style(style.NextStyleId, RtfStyleKind.Paragraph);
            RtfStyleKind linkedKind = style.Kind == RtfStyleKind.Paragraph ? RtfStyleKind.Character : RtfStyleKind.Paragraph;
            style.LinkedStyleId = map.Style(style.LinkedStyleId, linkedKind);
            if (style.ListId != 0) style.ListId = MapIndex(style.ListId, map.ListInstances);
            style.FontId = MapIndex(style.FontId, fonts);
            if (style.Kind == RtfStyleKind.Paragraph && !style.BasedOnStyleId.HasValue) style.FontId ??= map.DefaultFont;
            style.ForegroundColorIndex = MapIndex(style.ForegroundColorIndex, colors);
            style.HighlightColorIndex = MapIndex(style.HighlightColorIndex, colors);
            style.BackgroundColorIndex = MapIndex(style.BackgroundColorIndex, colors);
            style.ShadingForegroundColorIndex = MapIndex(style.ShadingForegroundColorIndex, colors);
            style.LegacyNumbering.FontId = MapIndex(style.LegacyNumbering.FontId, fonts);
            RemapBorder(style.TopBorder, colors);
            RemapBorder(style.LeftBorder, colors);
            RemapBorder(style.BottomBorder, colors);
            RemapBorder(style.RightBorder, colors);
            _styles.Add(style);
        }
        int? defaultStyle = map.Style(0, RtfStyleKind.Paragraph);
        if (!defaultStyle.HasValue) {
            var normal = new RtfStyle(styleId, "Imported Normal") {
                FontId = map.DefaultFont, Bold = false, Italic = false,
                UnderlineStyle = RtfUnderlineStyle.None, ParagraphAlignment = RtfTextAlignment.Left
            };
            _styles.Add(normal);
            defaultStyle = normal.Id;
        }
        map.DefaultParagraphStyle = defaultStyle.Value;
        return map;
    }
}
