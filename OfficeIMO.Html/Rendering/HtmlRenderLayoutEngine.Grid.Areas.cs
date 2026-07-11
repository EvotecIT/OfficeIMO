namespace OfficeIMO.Html;

internal sealed partial class HtmlRenderLayoutEngine {
    private IReadOnlyDictionary<string, GridAreaDefinition> ParseGridTemplateAreas(
        string value,
        string source,
        out int rowCount,
        out int columnCount) {
        rowCount = 0;
        columnCount = 0;
        if (string.IsNullOrWhiteSpace(value) || string.Equals(value.Trim(), "none", StringComparison.OrdinalIgnoreCase)) {
            return new Dictionary<string, GridAreaDefinition>(StringComparer.Ordinal);
        }

        List<string> rows = ExtractGridAreaRows(value);
        if (rows.Count == 0) {
            ReportUnsupportedGridValue(source, "grid-template-areas=" + value);
            return new Dictionary<string, GridAreaDefinition>(StringComparer.Ordinal);
        }
        if (rows.Count > _options.MaxGridTracks) EnsureGridPlacementLimit(rows.Count);

        var cells = new List<string[]>();
        foreach (string row in rows) {
            string[] names = row.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (names.Length == 0 || names.Length > _options.MaxGridTracks || columnCount > 0 && names.Length != columnCount) {
                ReportUnsupportedGridValue(source, "grid-template-areas=" + value);
                return new Dictionary<string, GridAreaDefinition>(StringComparer.Ordinal);
            }
            columnCount = names.Length;
            cells.Add(names);
        }

        rowCount = cells.Count;
        var bounds = new Dictionary<string, GridAreaDefinition>(StringComparer.Ordinal);
        for (int row = 0; row < cells.Count; row++) {
            for (int column = 0; column < cells[row].Length; column++) {
                string name = cells[row][column];
                if (name.All(character => character == '.')) continue;
                if (!bounds.TryGetValue(name, out GridAreaDefinition? area)) {
                    bounds[name] = new GridAreaDefinition(row, column, 1, 1);
                } else {
                    int rowEnd = Math.Max(area.Row + area.RowSpan, row + 1);
                    int columnEnd = Math.Max(area.Column + area.ColumnSpan, column + 1);
                    bounds[name] = new GridAreaDefinition(
                        Math.Min(area.Row, row),
                        Math.Min(area.Column, column),
                        rowEnd - Math.Min(area.Row, row),
                        columnEnd - Math.Min(area.Column, column));
                }
            }
        }

        foreach (KeyValuePair<string, GridAreaDefinition> pair in bounds.ToList()) {
            GridAreaDefinition area = pair.Value;
            bool rectangular = true;
            for (int row = area.Row; row < area.Row + area.RowSpan && rectangular; row++) {
                for (int column = area.Column; column < area.Column + area.ColumnSpan; column++) {
                    if (!string.Equals(cells[row][column], pair.Key, StringComparison.Ordinal)) {
                        rectangular = false;
                        break;
                    }
                }
            }
            if (!rectangular) {
                ReportUnsupportedGridValue(source, "grid-template-areas=" + pair.Key + " (non-rectangular)");
                bounds.Remove(pair.Key);
            }
        }
        return bounds;
    }

    private static List<string> ExtractGridAreaRows(string value) {
        var rows = new List<string>();
        char quote = '\0';
        int start = -1;
        for (int index = 0; index < value.Length; index++) {
            char current = value[index];
            if (quote == '\0') {
                if (current == '\'' || current == '"') {
                    quote = current;
                    start = index + 1;
                }
                continue;
            }
            if (current == quote && (index == 0 || value[index - 1] != '\\')) {
                rows.Add(value.Substring(start, index - start));
                quote = '\0';
                start = -1;
            }
        }
        return rows;
    }

    private sealed class GridAreaDefinition {
        internal GridAreaDefinition(int row, int column, int rowSpan, int columnSpan) {
            Row = row;
            Column = column;
            RowSpan = rowSpan;
            ColumnSpan = columnSpan;
        }
        internal int Row { get; }
        internal int Column { get; }
        internal int RowSpan { get; }
        internal int ColumnSpan { get; }
    }
}
