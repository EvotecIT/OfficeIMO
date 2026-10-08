using DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word {
    public partial class WordTable {
        /// <summary>Reads the named or default table style without materializing inherited properties in the source.</summary>
        private TableLayoutValues? GetInheritedTableLayoutType() {
            Styles? styles = _document._wordprocessingDocument?.MainDocumentPart?.StyleDefinitionsPart?.Styles;
            if (styles == null) return null;
            string? styleId = _tableProperties?.TableStyle?.Val?.Value;
            if (string.IsNullOrWhiteSpace(styleId)) {
                styleId = styles.Elements<Style>()
                    .FirstOrDefault(style => style.Type?.Value == StyleValues.Table && style.Default?.Value == true)
                    ?.StyleId?.Value;
            }
            var visited = new HashSet<string>(StringComparer.Ordinal);
            while (!string.IsNullOrWhiteSpace(styleId) && visited.Add(styleId!)) {
                // Word ignores child properties on the built-in Normal Table style.
                if (string.Equals(styleId, "TableNormal", StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(styleId, "NormalTable", StringComparison.OrdinalIgnoreCase)) return null;
                Style? style = styles.Elements<Style>().FirstOrDefault(candidate =>
                    candidate.Type?.Value == StyleValues.Table && candidate.StyleId?.Value == styleId);
                if (style == null) return null;
                TableLayoutValues? layout = style.GetFirstChild<StyleTableProperties>()?.GetFirstChild<TableLayout>()?.Type?.Value;
                if (layout.HasValue) return layout;
                styleId = style.BasedOn?.Val?.Value;
            }
            return null;
        }
    }
}
