using OfficeIMO.Rtf.Diagnostics;

namespace OfficeIMO.Reader.Rtf;

public static partial class DocumentReaderRtfExtensions {
    private static RtfReaderBlock CombineBlocks(IReadOnlyList<RtfReaderBlock> blocks) {
        string text = string.Join(Environment.NewLine + Environment.NewLine, blocks.Select(static block => block.Text).Where(static text => !string.IsNullOrWhiteSpace(text)));
        string markdown = string.Join(Environment.NewLine + Environment.NewLine, blocks.Select(static block => block.Markdown ?? block.Text).Where(static text => !string.IsNullOrWhiteSpace(text)));
        var tables = blocks.SelectMany(static block => block.Tables ?? Array.Empty<ReaderTable>()).ToArray();
        var visuals = blocks.SelectMany(static block => block.Visuals ?? Array.Empty<ReaderVisual>()).ToArray();
        var warnings = blocks.SelectMany(static block => block.Warnings ?? Array.Empty<string>()).ToArray();

        return new RtfReaderBlock(
            "document",
            0,
            text,
            markdown,
            tables.Length == 0 ? null : tables,
            visuals.Length == 0 ? null : visuals,
            warnings.Length == 0 ? null : warnings);
    }

    private static IReadOnlyList<string> SplitText(string text, int maxChars) {
        if (string.IsNullOrEmpty(text)) return Array.Empty<string>();
        if (text.Length <= maxChars) return new[] { text };

        var parts = new List<string>();
        int offset = 0;
        while (offset < text.Length) {
            int take = Math.Min(maxChars, text.Length - offset);
            parts.Add(text.Substring(offset, take));
            offset += take;
        }

        return parts;
    }

    private static IReadOnlyList<string>? BuildDiagnosticWarnings(IReadOnlyList<RtfDiagnostic> diagnostics) {
        if (diagnostics.Count == 0) return null;

        var warnings = new List<string>(diagnostics.Count);
        for (int i = 0; i < diagnostics.Count; i++) {
            RtfDiagnostic diagnostic = diagnostics[i];
            warnings.Add(diagnostic.Code + ": " + diagnostic.Message);
        }

        return warnings;
    }

    private static IReadOnlyList<string>? MergeWarnings(params object?[] sources) {
        List<string>? merged = null;
        for (int i = 0; i < sources.Length; i++) {
            switch (sources[i]) {
                case string value when !string.IsNullOrWhiteSpace(value):
                    merged ??= new List<string>();
                    AddUniqueWarning(merged, value);
                    break;
                case IEnumerable<string> values:
                    foreach (string item in values) {
                        if (string.IsNullOrWhiteSpace(item)) continue;
                        merged ??= new List<string>();
                        AddUniqueWarning(merged, item);
                    }
                    break;
            }
        }

        return merged;
    }

    private static void AddUniqueWarning(List<string> warnings, string warning) {
        for (int i = 0; i < warnings.Count; i++) {
            if (string.Equals(warnings[i], warning, StringComparison.Ordinal)) {
                return;
            }
        }

        warnings.Add(warning);
    }

    private static ReaderChunkDiagnostics BuildDiagnostics(RtfReaderBlock block, int documentLinkCount, int documentFormFieldCount) {
        int imageCount = block.Visuals?.Count(static visual => string.Equals(visual.Kind, "image", StringComparison.OrdinalIgnoreCase)) ?? 0;
        return new ReaderChunkDiagnostics {
            SourceKind = "rtf",
            TableCount = block.Tables?.Count ?? 0,
            ImageCount = imageCount,
            SelectedPageCount = 0,
            PageCount = 0,
            LinkCount = documentLinkCount,
            FormFieldCount = documentFormFieldCount
        };
    }

    private static int CountHyperlinkRuns(RtfDocument document) {
        int count = 0;
        for (int i = 0; i < document.Blocks.Count; i++) {
            count += CountHyperlinkRuns(document.Blocks[i]);
        }

        for (int i = 0; i < document.HeaderFooters.Count; i++) {
            RtfHeaderFooter headerFooter = document.HeaderFooters[i];
            for (int paragraphIndex = 0; paragraphIndex < headerFooter.Paragraphs.Count; paragraphIndex++) {
                count += CountHyperlinkRuns(headerFooter.Paragraphs[paragraphIndex]);
            }
        }

        for (int i = 0; i < document.Notes.Count; i++) {
            RtfNote note = document.Notes[i];
            for (int paragraphIndex = 0; paragraphIndex < note.Paragraphs.Count; paragraphIndex++) {
                count += CountHyperlinkRuns(note.Paragraphs[paragraphIndex]);
            }
        }

        return count;
    }

    private static int CountHyperlinkRuns(IRtfBlock block) {
        switch (block) {
            case RtfParagraph paragraph:
                return CountHyperlinkRuns(paragraph);
            case RtfTable table:
                return CountHyperlinkRuns(table);
            case RtfObject rtfObject:
                return CountHyperlinkRuns(rtfObject.Result);
            case RtfShape shape:
                return CountHyperlinkRuns(shape);
            default:
                return 0;
        }
    }

    private static int CountHyperlinkRuns(RtfTable table) {
        int count = 0;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            RtfTableRow row = table.Rows[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
                RtfTableCell cell = row.Cells[cellIndex];
                for (int blockIndex = 0; blockIndex < cell.Blocks.Count; blockIndex++) {
                    count += CountHyperlinkRuns(cell.Blocks[blockIndex]);
                }
            }
        }

        return count;
    }

    private static int CountHyperlinkRuns(RtfShape shape) {
        int count = 0;
        for (int paragraphIndex = 0; paragraphIndex < shape.TextBoxParagraphs.Count; paragraphIndex++) {
            count += CountHyperlinkRuns(shape.TextBoxParagraphs[paragraphIndex]);
        }

        return count;
    }

    private static int CountHyperlinkRuns(RtfParagraph paragraph) {
        int count = 0;
        for (int runIndex = 0; runIndex < paragraph.Runs.Count; runIndex++) {
            if (paragraph.Runs[runIndex].Hyperlink != null) count++;
        }

        if (paragraph.ListText != null) {
            count += CountHyperlinkRuns(paragraph.ListText);
        }

        for (int inlineIndex = 0; inlineIndex < paragraph.Inlines.Count; inlineIndex++) {
            switch (paragraph.Inlines[inlineIndex]) {
                case RtfField field:
                    count += CountHyperlinkRuns(field.Result);
                    break;
                case RtfObject rtfObject:
                    count += CountHyperlinkRuns(rtfObject.Result);
                    break;
                case RtfShape shape:
                    count += CountHyperlinkRuns(shape);
                    break;
            }
        }

        return count;
    }

    private static int CountFormFields(RtfDocument document) {
        int count = 0;
        for (int i = 0; i < document.Blocks.Count; i++) {
            count += CountFormFields(document.Blocks[i]);
        }

        for (int i = 0; i < document.HeaderFooters.Count; i++) {
            RtfHeaderFooter headerFooter = document.HeaderFooters[i];
            for (int paragraphIndex = 0; paragraphIndex < headerFooter.Paragraphs.Count; paragraphIndex++) {
                count += CountFormFields(headerFooter.Paragraphs[paragraphIndex]);
            }
        }

        for (int i = 0; i < document.Notes.Count; i++) {
            RtfNote note = document.Notes[i];
            for (int paragraphIndex = 0; paragraphIndex < note.Paragraphs.Count; paragraphIndex++) {
                count += CountFormFields(note.Paragraphs[paragraphIndex]);
            }
        }

        return count;
    }

    private static int CountFormFields(IRtfBlock block) {
        switch (block) {
            case RtfParagraph paragraph:
                return CountFormFields(paragraph);
            case RtfTable table:
                return CountFormFields(table);
            case RtfObject rtfObject:
                return CountFormFields(rtfObject.Result);
            case RtfShape shape:
                return CountFormFields(shape);
            default:
                return 0;
        }
    }

    private static int CountFormFields(RtfTable table) {
        int count = 0;
        for (int rowIndex = 0; rowIndex < table.Rows.Count; rowIndex++) {
            RtfTableRow row = table.Rows[rowIndex];
            for (int cellIndex = 0; cellIndex < row.Cells.Count; cellIndex++) {
                RtfTableCell cell = row.Cells[cellIndex];
                for (int blockIndex = 0; blockIndex < cell.Blocks.Count; blockIndex++) {
                    count += CountFormFields(cell.Blocks[blockIndex]);
                }
            }
        }

        return count;
    }

    private static int CountFormFields(RtfShape shape) {
        int count = 0;
        for (int paragraphIndex = 0; paragraphIndex < shape.TextBoxParagraphs.Count; paragraphIndex++) {
            count += CountFormFields(shape.TextBoxParagraphs[paragraphIndex]);
        }

        return count;
    }

    private static int CountFormFields(RtfParagraph paragraph) {
        int count = 0;
        if (paragraph.ListText != null) {
            count += CountFormFields(paragraph.ListText);
        }

        for (int inlineIndex = 0; inlineIndex < paragraph.Inlines.Count; inlineIndex++) {
            switch (paragraph.Inlines[inlineIndex]) {
                case RtfField field:
                    if (field.FormFieldData != null) count++;
                    count += CountFormFields(field.Result);
                    break;
                case RtfObject rtfObject:
                    count += CountFormFields(rtfObject.Result);
                    break;
                case RtfShape shape:
                    count += CountFormFields(shape);
                    break;
            }
        }

        return count;
    }

    private static string BuildChunkId(string kind, int sourceBlockIndex, int partIndex, int partCount) {
        string baseId = "rtf-" + kind + "-" + sourceBlockIndex.ToString("D4", CultureInfo.InvariantCulture);
        return partCount <= 1 ? baseId : baseId + "-part-" + partIndex.ToString("D4", CultureInfo.InvariantCulture);
    }

    private static string BuildBlockAnchor(string kind, int sourceBlockIndex, int partIndex, int partCount) {
        string anchor = "rtf-" + kind + "-" + sourceBlockIndex.ToString("D4", CultureInfo.InvariantCulture);
        return partCount <= 1 ? anchor : anchor + "-part-" + partIndex.ToString("D4", CultureInfo.InvariantCulture);
    }

    private sealed class RtfReaderBlock {
        public RtfReaderBlock(string kind, int sourceBlockIndex, string text, string? markdown, IReadOnlyList<ReaderTable>? tables, IReadOnlyList<ReaderVisual>? visuals, IReadOnlyList<string>? warnings) {
            Kind = kind;
            SourceBlockIndex = sourceBlockIndex;
            Text = text ?? string.Empty;
            Markdown = markdown;
            Tables = tables;
            Visuals = visuals;
            Warnings = warnings;
        }

        public string Kind { get; }
        public int SourceBlockIndex { get; }
        public string Text { get; }
        public string? Markdown { get; }
        public IReadOnlyList<ReaderTable>? Tables { get; }
        public IReadOnlyList<ReaderVisual>? Visuals { get; }
        public IReadOnlyList<string>? Warnings { get; }
    }
}
