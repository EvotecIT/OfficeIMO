namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class TableParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.Tables) return false;
            if (IsHeaderlessSingleRowTableMarker(lines[i])) {
                return TryParseHeaderlessSingleRowTable(lines, ref i, options, doc, state);
            }

            if (!LooksLikeTableRow(lines[i])) return false;
            if (!TryGetTableExtent(
                lines,
                i,
                out int end,
                out _,
                allowHeaderlessTables: options.AllowHeaderlessTables,
                options: options,
                allowMismatchedAlignmentCells: state.IsMarkdigDefinitionListBody)) return false;

            var table = ParseTable(lines, i, end, options, state);
            doc.Add(table); i = end + 1; return true;
        }
    }

    private static bool IsHeaderlessSingleRowTableMarker(string? line) {
        return string.Equals((line ?? string.Empty).Trim(), TableBlock.HeaderlessSingleRowTableMarker, StringComparison.Ordinal);
    }

    private static bool TryParseHeaderlessSingleRowTable(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
        int tableStart = i + 1;
        if (tableStart >= lines.Length || !LooksLikeTableRow(lines[tableStart])) {
            return false;
        }

        if (!TryGetTableExtent(lines, tableStart, out int end, out _, allowSingleRowHeaderless: true, options: options)) {
            return false;
        }

        TableBlock table = ParseTable(lines, tableStart, end, options, state);
        if (table.Headers.Count == 0 && table.Rows.Count == 1) {
            table.PreserveHeaderlessSingleRowTable = true;
        }

        doc.Add(table);
        i = end + 1;
        return true;
    }
}
