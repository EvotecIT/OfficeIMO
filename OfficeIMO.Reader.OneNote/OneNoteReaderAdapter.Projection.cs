using OfficeIMO.OneNote;
using OfficeIMO.OneNote.Markdown;

namespace OfficeIMO.Reader.OneNote;

internal static partial class OneNoteReaderAdapter {
    private static IEnumerable<ReaderChunk> BuildChunks(
        OneNoteSection section,
        SourceInfo source,
        ReaderOptions options,
        CancellationToken cancellationToken,
        IReadOnlyList<string>? hierarchyOverride = null) {
        IReadOnlyList<string> hierarchy = hierarchyOverride ?? BuildPageHierarchy(section);
        if (hierarchy.Count != section.Pages.Count) throw new ArgumentException("The OneNote page hierarchy count must match the page count.", nameof(hierarchyOverride));
        for (int pageIndex = 0; pageIndex < section.Pages.Count; pageIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            OneNotePage page = section.Pages[pageIndex];
            string text = BuildPageText(page);
            string markdown = BuildPageMarkdown(page, pageIndex);
            ReaderTable[] tables = BuildTables(page, source, pageIndex, options.MaxTableRows).ToArray();
            string[] warnings = section.Diagnostics.Concat(page.Diagnostics)
                .Where(static diagnostic => diagnostic.Severity != OneNoteDiagnosticSeverity.Information)
                .Select(static diagnostic => diagnostic.Code + ": " + diagnostic.Message)
                .Distinct(StringComparer.Ordinal)
                .ToArray();

            IReadOnlyList<string> textParts = SplitByMaxChars(text, options.MaxChars);
            IReadOnlyList<string> markdownParts = SplitByMaxChars(markdown, options.MaxChars);
            int partCount = Math.Max(textParts.Count, markdownParts.Count);
            for (int partIndex = 0; partIndex < partCount; partIndex++) {
                string anchor = "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) +
                    (partCount == 1 ? string.Empty : "-part-" + (partIndex + 1).ToString(CultureInfo.InvariantCulture));
                var chunk = new ReaderChunk {
                    Id = "onenote-page-" + (pageIndex + 1).ToString("D4", CultureInfo.InvariantCulture) +
                        (partCount == 1 ? string.Empty : "-part-" + (partIndex + 1).ToString("D4", CultureInfo.InvariantCulture)),
                    Kind = ReaderInputKind.OneNote,
                    Location = BuildLocation(source, pageIndex, "page", anchor, hierarchy[pageIndex]),
                    SourceId = source.SourceId,
                    SourceHash = source.SourceHash,
                    SourceLastWriteUtc = source.LastWriteUtc,
                    SourceLengthBytes = source.LengthBytes,
                    Text = partIndex < textParts.Count ? textParts[partIndex] : string.Empty,
                    Markdown = partIndex < markdownParts.Count ? markdownParts[partIndex] : string.Empty,
                    Tables = partIndex == 0 && tables.Length > 0 ? tables : null,
                    Warnings = partIndex == 0 && warnings.Length > 0 ? warnings : null
                };
                chunk.TokenEstimate = EstimateTokenCount(chunk.Markdown ?? chunk.Text);
                if (options.ComputeHashes) chunk.ChunkHash = ComputeHash(BuildChunkHashInput(chunk));
                yield return chunk;
            }
        }
    }

    private static string[] BuildPageHierarchy(OneNoteSection section) {
        var result = new string[section.Pages.Count];
        var stack = new List<string>();
        for (int index = 0; index < section.Pages.Count; index++) {
            OneNotePage page = section.Pages[index];
            int level = Math.Max(0, page.Level);
            while (stack.Count > level) stack.RemoveAt(stack.Count - 1);
            while (stack.Count < level) stack.Add("Untitled");
            string title = string.IsNullOrWhiteSpace(page.Title) ? "Untitled page" : page.Title;
            if (stack.Count == level) stack.Add(title); else stack[level] = title;
            result[index] = string.Join(" > ", new[] { section.Name }.Concat(stack));
            if (stack.Count > level + 1) stack.RemoveRange(level + 1, stack.Count - level - 1);
        }
        return result;
    }

    private static string BuildPageText(OneNotePage page) {
        return OneNoteMarkdownProjection.ToText(page);
    }

    private static string BuildPageMarkdown(OneNotePage page, int pageIndex) {
        int assetIndex = 0;
        int headingLevel = Math.Min(6, Math.Max(1, page.Level + 1));
        return OneNoteMarkdownProjection.ToMarkdown(
            page,
            headingLevel,
            _ => BuildAssetId(pageIndex, assetIndex++));
    }

    private static IEnumerable<ReaderTable> BuildTables(OneNotePage page, SourceInfo source, int pageIndex, int maxRows) {
        int tableIndex = 0;
        foreach (OneNoteElement element in EnumerateAllElements(page)) {
            if (!(element is OneNoteTable table)) continue;
            int columns = table.Rows.Count == 0 ? table.ColumnWidths.Count : table.Rows.Max(static row => row.Cells.Count);
            string[] headers = Enumerable.Range(1, columns).Select(static index => "Column " + index.ToString(CultureInfo.InvariantCulture)).ToArray();
            IReadOnlyList<IReadOnlyList<string>> rows = table.Rows.Select(row =>
                (IReadOnlyList<string>)Enumerable.Range(0, columns)
                    .Select(column => column < row.Cells.Count ? CellText(row.Cells[column]) : string.Empty)
                    .ToArray()).ToArray();
            IReadOnlyList<IReadOnlyList<string>> visible = rows.Take(Math.Max(0, maxRows)).ToArray();
            yield return new ReaderTable {
                Title = (string.IsNullOrWhiteSpace(page.Title) ? "OneNote page" : page.Title) + " table " + (tableIndex + 1).ToString(CultureInfo.InvariantCulture),
                Kind = "onenote-table",
                Location = BuildLocation(source, pageIndex, "table", "page-" + (pageIndex + 1).ToString(CultureInfo.InvariantCulture) + "-table-" + (tableIndex + 1).ToString(CultureInfo.InvariantCulture), null, tableIndex),
                Columns = headers,
                ColumnProfiles = ReaderTableProfiler.CreateProfiles(headers, visible),
                Rows = visible,
                TotalRowCount = rows.Count,
                Truncated = rows.Count > visible.Count
            };
            tableIndex++;
        }
    }

    private static string CellText(OneNoteTableCell cell) {
        return OneNoteMarkdownProjection.ToText(cell);
    }

    private static IEnumerable<OneNoteElement> EnumeratePageRoots(OneNotePage page) {
        foreach (OneNoteOutline outline in page.Outlines) yield return outline;
        foreach (OneNoteElement element in page.DirectContent) yield return element;
    }

    private static IEnumerable<OneNoteElement> EnumerateAllElements(OneNotePage page) {
        foreach (OneNoteElement root in EnumeratePageRoots(page)) {
            foreach (OneNoteElement element in EnumerateElementTree(root)) yield return element;
        }
    }

    private static IEnumerable<OneNoteElement> EnumerateElementTree(OneNoteElement element) {
        yield return element;
        if (element is OneNoteOutline outline) {
            foreach (OneNoteElement child in outline.Children)
                foreach (OneNoteElement nested in EnumerateElementTree(child)) yield return nested;
        } else if (element is OneNoteParagraph paragraph) {
            foreach (OneNoteElement child in paragraph.Children)
                foreach (OneNoteElement nested in EnumerateElementTree(child)) yield return nested;
        } else if (element is OneNoteTable table) {
            foreach (OneNoteTableRow row in table.Rows)
                foreach (OneNoteTableCell cell in row.Cells)
                    foreach (OneNoteElement child in cell.Content)
                        foreach (OneNoteElement nested in EnumerateElementTree(child)) yield return nested;
        }
    }

    private static ReaderLocation BuildLocation(SourceInfo source, int pageIndex, string blockKind, string anchor, string? hierarchy = null, int? tableIndex = null) {
        return new ReaderLocation {
            Path = source.Path,
            Page = pageIndex + 1,
            SourceBlockIndex = pageIndex,
            SourceBlockKind = blockKind,
            BlockAnchor = anchor,
            HeadingPath = hierarchy,
            HierarchyHeadingPath = hierarchy,
            TableIndex = tableIndex
        };
    }

    private static IReadOnlyList<string> SplitByMaxChars(string value, int maxChars) {
        maxChars = Math.Max(1, maxChars);
        if (string.IsNullOrEmpty(value)) return new[] { string.Empty };
        var parts = new List<string>();
        int offset = 0;
        while (offset < value.Length) {
            int length = Math.Min(maxChars, value.Length - offset);
            int cut = length;
            if (offset + length < value.Length) {
                int newline = value.LastIndexOf('\n', offset + length - 1, length);
                if (newline > offset) cut = newline - offset + 1;
                else {
                    int space = value.LastIndexOf(' ', offset + length - 1, length);
                    if (space > offset) cut = space - offset + 1;
                }
            }
            string part = value.Substring(offset, cut).Trim();
            if (part.Length == 0) part = value.Substring(offset, cut);
            parts.Add(part);
            offset += cut;
            while (offset < value.Length && char.IsWhiteSpace(value[offset])) offset++;
        }
        return parts;
    }

    private static string EscapeMarkdown(string? value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        return value!.Replace("\\", "\\\\").Replace("*", "\\*").Replace("_", "\\_").Replace("[", "\\[").Replace("]", "\\]").Replace("|", "\\|");
    }

}
