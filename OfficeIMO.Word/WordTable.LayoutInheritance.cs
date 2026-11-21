using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTable {
        private const int MaxTableLayoutStyleDepth = 256;

        /// <summary>Reads the named or default table style without materializing inherited properties in the source.</summary>
        private TableLayoutValues? GetInheritedTableLayoutType() {
            var definitions = new Dictionary<string, WordStyleDefinitionsReader.StyleDefinition>(StringComparer.Ordinal);
            string? defaultStyleId = null;
            foreach (WordStyleDefinitionsReader.StyleDefinition style in WordStyleCatalog.For(_document).Definitions) {
                if (!string.Equals(style.Type, "table", StringComparison.Ordinal) || string.IsNullOrWhiteSpace(style.StyleId)) continue;
                string id = style.StyleId!;
                if (!definitions.ContainsKey(id)) definitions.Add(id, style);
                if (defaultStyleId == null && style.IsDefault) defaultStyleId = id;
            }
            string? styleId = _tableProperties?.TableStyle?.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) styleId = defaultStyleId;
            var visited = new HashSet<string>(StringComparer.Ordinal);
            while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                if (visited.Count > MaxTableLayoutStyleDepth) throw new InvalidDataException("The Word table style inheritance chain exceeds the supported depth.");
                // Word ignores child properties on the built-in Normal Table style.
                if (string.Equals(styleId, "TableNormal", StringComparison.Ordinal) ||
                    string.Equals(styleId, "NormalTable", StringComparison.Ordinal)) return null;
                if (!definitions.TryGetValue(styleId!, out WordStyleDefinitionsReader.StyleDefinition? style)) return null;
                TableLayoutValues? layout = style.Element.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableLayout>()?.Type?.Value;
                if (layout.HasValue) return layout;
                styleId = style.BasedOn;
            }
            return null;
        }
    }
}
