using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentReaderAdapter {
    private static IEnumerable<ReaderChunk> ReadTextDocument(OdtDocument document, string sourceName, ReaderOptions options,
        CancellationToken cancellationToken) {
        var headings = new string?[10];
        int blockIndex = 0;
        foreach (OdtContentBlock block in document.ContentBlocks) {
            cancellationToken.ThrowIfCancellationRequested();
            if (block.Table != null) {
                yield return BuildTableChunk(block.Table, sourceName, blockIndex, options, cancellationToken);
                blockIndex++;
                continue;
            }

            OdtParagraph paragraph = block.Paragraph!;
            string text = paragraph.Text.Trim();
            int headingLevel = Math.Max(1, Math.Min(10, paragraph.HeadingLevel ?? 1));
            if (paragraph.IsHeading) {
                headings[headingLevel - 1] = text;
                for (int index = headingLevel; index < headings.Length; index++) headings[index] = null;
            }
            string? headingPath = string.Join(" > ", headings.Where(value => !string.IsNullOrWhiteSpace(value))!);
            int part = 0;
            foreach (string piece in SplitText(text, options.MaxChars)) {
                string markdown = paragraph.IsHeading
                    ? new string('#', Math.Min(6, headingLevel)) + " " + piece
                    : piece;
                yield return new ReaderChunk {
                    Id = BuildId(sourceName, paragraph.IsHeading ? "heading" : "paragraph", blockIndex, part++),
                    Kind = ReaderInputKind.OpenDocument,
                    Location = new ReaderLocation {
                        Path = sourceName,
                        BlockIndex = blockIndex,
                        SourceBlockIndex = blockIndex,
                        SourceBlockKind = paragraph.IsHeading ? "heading" : "paragraph",
                        HeadingPath = headingPath
                    },
                    Text = piece,
                    Markdown = markdown
                };
            }
            blockIndex++;
        }
    }

    private static ReaderChunk BuildTableChunk(OdtTable table, string sourceName, int blockIndex, ReaderOptions options,
        CancellationToken cancellationToken) {
        IReadOnlyList<OdtTableRow> rows = table.Rows;
        int maximumRows = options.MaxTableRows > 0 ? options.MaxTableRows : 200;
        OdtTableRow[] selectedRows = rows.Take(maximumRows).ToArray();
        int columnCount = selectedRows.Length == 0
            ? 0
            : Math.Min(MaximumTableColumns, selectedRows.Max(row => row.Cells.Count));
        string[] columns = Enumerable.Range(1, columnCount).Select(index => "Column " + index.ToString(CultureInfo.InvariantCulture)).ToArray();
        var values = new List<IReadOnlyList<string>>(selectedRows.Length);
        foreach (OdtTableRow row in selectedRows) {
            cancellationToken.ThrowIfCancellationRequested();
            values.Add(Enumerable.Range(0, columnCount)
                .Select(index => index < row.Cells.Count ? row.Cells[index].Text : string.Empty).ToArray());
        }
        var readerTable = new ReaderTable {
            Title = table.Name,
            Kind = "odt-table",
            Columns = columns,
            Rows = values,
            TotalRowCount = rows.Count,
            Truncated = rows.Count > maximumRows || selectedRows.Any(row => row.Cells.Count > MaximumTableColumns),
            Location = new ReaderLocation { Path = sourceName, BlockIndex = blockIndex, SourceBlockIndex = blockIndex, SourceBlockKind = "table" }
        };
        string text = string.Join(Environment.NewLine, values.Select(row => string.Join("\t", row)));
        string markdown = BuildTableMarkdown(columns, values);
        return new ReaderChunk {
            Id = BuildId(sourceName, "table", blockIndex),
            Kind = ReaderInputKind.OpenDocument,
            Location = readerTable.Location!,
            Text = text,
            Markdown = markdown,
            Tables = new[] { readerTable },
            Warnings = readerTable.Truncated ? new[] { "Table rows were truncated due to MaxTableRows." } : null
        };
    }

    private static string BuildTableMarkdown(IReadOnlyList<string> columns, IReadOnlyList<IReadOnlyList<string>> rows) {
        if (columns.Count == 0) return string.Empty;
        string Header(IEnumerable<string> cells) => "| " + string.Join(" | ", cells.Select(value => value.Replace("|", "\\|"))) + " |";
        var lines = new List<string> { Header(columns), Header(columns.Select(_ => "---")) };
        lines.AddRange(rows.Select(Header));
        return string.Join(Environment.NewLine, lines);
    }
}
