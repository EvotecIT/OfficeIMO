using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTable {
        private const int MaxTableLayoutStyleDepth = 256;

        /// <summary>Reads the named or default table style without materializing inherited properties in the source.</summary>
        private TableLayoutValues? GetInheritedTableLayoutType() {
            Styles? styles = _document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) return null;
            var definitions = new Dictionary<string, Style>(StringComparer.Ordinal);
            string? defaultStyleId = null;
            foreach (Style style in styles.Elements<Style>()) {
                if (style.Type?.Value != StyleValues.Table || string.IsNullOrWhiteSpace(style.StyleId?.Value)) continue;
                string id = style.StyleId!.Value!;
                if (!definitions.ContainsKey(id)) definitions.Add(id, style);
                if (defaultStyleId == null && style.Default?.Value == true) defaultStyleId = id;
            }
            string? styleId = _tableProperties?.TableStyle?.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) styleId = defaultStyleId;
            var visited = new HashSet<string>(StringComparer.Ordinal);
            while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                if (visited.Count > MaxTableLayoutStyleDepth) throw new InvalidDataException("The Word table style inheritance chain exceeds the supported depth.");
                // Word ignores child properties on the built-in Normal Table style.
                if (string.Equals(styleId, "TableNormal", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(styleId, "NormalTable", StringComparison.OrdinalIgnoreCase)) return null;
                if (!definitions.TryGetValue(styleId!, out Style? style)) return null;
                TableLayoutValues? layout = style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableLayout>()?.Type?.Value;
                if (layout.HasValue) return layout;
                styleId = style.BasedOn?.Val?.Value;
            }
            return null;
        }
    }
}
