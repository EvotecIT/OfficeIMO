using OfficeIMO.OpenDocument;

namespace OfficeIMO.Reader.OpenDocument;

internal static partial class OpenDocumentReaderAdapter {
    private static IEnumerable<ReaderChunk> ReadPresentation(OdpPresentation document, string sourceName, ReaderOptions options, ReaderOpenDocumentOptions formatOptions,
        CancellationToken cancellationToken) {
        for (int slideIndex = 0; slideIndex < document.Slides.Count; slideIndex++) {
            cancellationToken.ThrowIfCancellationRequested();
            OdpSlide slide = document.Slides[slideIndex];
            var paragraphs = new List<string>();
            var tables = new List<ReaderTable>();
            var warnings = new List<string>();
            CollectSlideContent(slide.Shapes, sourceName, slideIndex, options, paragraphs, tables, warnings, cancellationToken);
            var notes = formatOptions.IncludeSpeakerNotes
                ? slide.SpeakerNotes?.Paragraphs.Select(paragraph => paragraph.Text.Trim()).Where(text => text.Length > 0).ToArray()
                : null;
            if (notes != null && notes.Length > 0) paragraphs.AddRange(notes.Select(text => "Notes: " + text));
            string text = string.Join(Environment.NewLine, paragraphs);
            var markdown = new StringBuilder();
            markdown.Append("## Slide ").Append(slideIndex + 1).Append(": ").AppendLine(slide.Name);
            if (paragraphs.Count > 0) markdown.AppendLine().AppendLine(string.Join(Environment.NewLine + Environment.NewLine, paragraphs));
            foreach (ReaderTable table in tables) {
                markdown.AppendLine().AppendLine(BuildTableMarkdown(table.Columns, table.Rows));
            }
            if (slide.Hidden) warnings.Add("Slide is hidden in the source presentation.");
            if (!formatOptions.IncludeSpeakerNotes && slide.SpeakerNotes != null) warnings.Add("Speaker notes were omitted by ReaderOpenDocumentOptions.");
            yield return new ReaderChunk {
                Id = BuildId(sourceName, "slide", slideIndex), Kind = ReaderInputKind.OpenDocument,
                Location = new ReaderLocation {
                    Path = sourceName, BlockIndex = slideIndex, SourceBlockIndex = slideIndex,
                    SourceBlockKind = "slide", Slide = slideIndex + 1, HeadingPath = slide.Name
                },
                Text = text, Markdown = markdown.ToString().TrimEnd(), Tables = tables.Count == 0 ? null : tables,
                Warnings = warnings.Count == 0 ? null : warnings
            };
        }
    }

    private static void CollectSlideContent(IEnumerable<OdpShape> shapes, string sourceName, int slideIndex, ReaderOptions options,
        List<string> paragraphs, List<ReaderTable> tables, List<string> warnings, CancellationToken cancellationToken) {
        foreach (OdpShape shape in shapes) {
            cancellationToken.ThrowIfCancellationRequested();
            if (shape is OdpTextBox textBox) {
                paragraphs.AddRange(textBox.Paragraphs.Select(paragraph => paragraph.Text.Trim()).Where(text => text.Length > 0));
            } else if (shape is OdpTable table) {
                tables.Add(BuildPresentationTable(table, sourceName, slideIndex, tables.Count, options, cancellationToken,
                    out bool rowsTruncated, out bool columnsTruncated));
                if (rowsTruncated && !warnings.Contains("Table rows were truncated due to MaxTableRows.")) {
                    warnings.Add("Table rows were truncated due to MaxTableRows.");
                }
                if (columnsTruncated && !warnings.Contains("Table columns were truncated to 256 columns for bounded extraction.")) {
                    warnings.Add("Table columns were truncated to 256 columns for bounded extraction.");
                }
            } else if (shape is OdpGroup group) {
                CollectSlideContent(group.Shapes, sourceName, slideIndex, options, paragraphs, tables, warnings, cancellationToken);
            }
        }
    }

    private static ReaderTable BuildPresentationTable(OdpTable table, string sourceName, int slideIndex, int tableIndex, ReaderOptions options,
        CancellationToken cancellationToken, out bool rowsTruncated, out bool columnsTruncated) {
        int maxRows = options.MaxTableRows > 0 ? options.MaxTableRows : 200;
        IReadOnlyList<OdpTableRow> sourceRows = table.Rows;
        OdpTableRow[] selectedRows = sourceRows.Take(maxRows).ToArray();
        rowsTruncated = sourceRows.Count > maxRows;
        columnsTruncated = selectedRows.Any(row => row.Cells.Count > MaximumTableColumns);
        int columnCount = selectedRows.Length == 0
            ? 0
            : Math.Min(MaximumTableColumns, selectedRows.Max(row => row.Cells.Count));
        string[] columns = Enumerable.Range(1, columnCount).Select(index => "Column " + index.ToString(CultureInfo.InvariantCulture)).ToArray();
        var rows = new List<IReadOnlyList<string>>(selectedRows.Length);
        foreach (OdpTableRow row in selectedRows) {
            cancellationToken.ThrowIfCancellationRequested();
            rows.Add(Enumerable.Range(0, columnCount)
                .Select(index => index < row.Cells.Count ? row.Cells[index].Text : string.Empty).ToArray());
        }
        return new ReaderTable {
            Title = table.Name, Kind = "odp-table", Columns = columns, Rows = rows,
            TotalRowCount = sourceRows.Count,
            Truncated = rowsTruncated || columnsTruncated,
            Location = new ReaderLocation { Path = sourceName, Slide = slideIndex + 1, TableIndex = tableIndex, SourceBlockKind = "table" }
        };
    }
}
