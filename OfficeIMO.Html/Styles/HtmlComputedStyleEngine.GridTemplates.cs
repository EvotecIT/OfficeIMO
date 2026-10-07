using AngleSharp.Css.Parser;
using System.Text;

namespace OfficeIMO.Html;

public static partial class HtmlComputedStyleEngine {
    private static bool IsGridTemplateSyntax(string propertyName, string value) {
        string normalized = value.Trim();
        if (propertyName != "grid-template-areas" && normalized.StartsWith("subgrid", StringComparison.OrdinalIgnoreCase)) {
            return IsSubgridTemplateSyntax(normalized);
        }

        // Keep the existing parser as the grammar owner. Preserve only the native
        // length-math extension that it cannot parse, then validate the surrounding
        // repeat/minmax/track and line-name grammar through the same parser.
        var parser = new CssParser();
        string? parsedValue = parser.ParseDeclaration(propertyName + ":" + value)?.GetPropertyValue(propertyName);
        if (!string.IsNullOrEmpty(parsedValue)) {
            return propertyName != "grid-template-areas" || AreGridTemplateAreasRectangular(parsedValue!);
        }
        return propertyName != "grid-template-areas"
            && TryNormalizeGridTemplateLengthMath(normalized, out string parserValue)
            && !string.IsNullOrEmpty(parser.ParseDeclaration(propertyName + ":" + parserValue)?.GetPropertyValue(propertyName));
    }

    private static bool AreGridTemplateAreasRectangular(string value) {
        if (string.Equals(value, "none", StringComparison.OrdinalIgnoreCase)) return true;
        IReadOnlyList<string> rows = HtmlRenderCssValues.SplitWhitespace(value);
        var bounds = new Dictionary<string, (int FirstRow, int LastRow, int FirstColumn, int LastColumn, int Count)>(StringComparer.Ordinal);
        int columnCount = 0;
        for (int row = 0; row < rows.Count; row++) {
            string text = rows[row];
            if (text.Length < 2 || text[0] is not ('\'' or '"') || text[text.Length - 1] != text[0]) return false;
            string[] cells = text.Substring(1, text.Length - 2).Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (cells.Length == 0 || row > 0 && cells.Length != columnCount) return false;
            columnCount = cells.Length;
            for (int column = 0; column < cells.Length; column++) {
                string name = cells[column];
                if (name.All(character => character == '.')) continue;
                bounds[name] = bounds.TryGetValue(name, out var area)
                    ? (area.FirstRow, row, Math.Min(area.FirstColumn, column), Math.Max(area.LastColumn, column), area.Count + 1)
                    : (row, row, column, column, 1);
            }
        }
        // Every occurrence is inside its bounds; equal cell and rectangle areas
        // prove that there are no holes or disconnected pieces.
        return rows.Count > 0 && bounds.Values.All(area =>
            (long)(area.LastRow - area.FirstRow + 1) * (area.LastColumn - area.FirstColumn + 1) == area.Count);
    }

    private static bool IsSubgridTemplateSyntax(string value) {
        int position = "subgrid".Length;
        while (position < value.Length) {
            while (position < value.Length && char.IsWhiteSpace(value[position])) position++;
            if (position == value.Length) break;
            if (value[position++] != '[') return false;
            int close = value.IndexOf(']', position);
            if (close < 0) return false;
            string names = value.Substring(position, close - position);
            int namePosition = 0;
            while (namePosition < names.Length) {
                while (namePosition < names.Length && char.IsWhiteSpace(names[namePosition])) namePosition++;
                if (namePosition == names.Length) break;
                if (!HtmlCssIdentifierParser.TryRead(names, ref namePosition, out string name) || !IsGridCustomIdentifier(name)
                    || namePosition < names.Length && !char.IsWhiteSpace(names[namePosition])) return false;
            }
            position = close + 1;
        }
        return true;
    }

    private static bool TryNormalizeGridTemplateLengthMath(string value, out string normalized) {
        var result = new StringBuilder(value.Length);
        bool changed = false;
        int copied = 0;
        for (int index = 0; index < value.Length; index++) {
            if (!char.IsLetter(value[index]) || index > 0 && (char.IsLetterOrDigit(value[index - 1]) || value[index - 1] is '-' or '_')) continue;
            int endName = index;
            if (!HtmlCssIdentifierParser.TryRead(value, ref endName, out string name)
                || endName >= value.Length || value[endName] != '(') continue;
            if (name.ToLowerInvariant() is not ("calc" or "min" or "max" or "clamp")) continue;
            int close = HtmlRenderCssValues.FindMatchingParenthesis(value, endName);
            if (close < 0 || !HtmlRenderCssValues.TryLength(value.Substring(index, close - index + 1),
                100D, 16D, 16D, 100D, 100D, 100D, 100D, out _)) {
                normalized = string.Empty;
                return false;
            }
            result.Append(value, copied, index - copied).Append("0px");
            copied = close + 1;
            index = close;
            changed = true;
        }
        result.Append(value, copied, value.Length - copied);
        normalized = result.ToString();
        return changed;
    }
}
