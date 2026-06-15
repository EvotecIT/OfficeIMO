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

    private static ReaderChunkDiagnostics BuildDiagnostics(RtfDocument document, RtfReaderBlock block) {
        int imageCount = block.Visuals?.Count(static visual => string.Equals(visual.Kind, "image", StringComparison.OrdinalIgnoreCase)) ?? 0;
        return new ReaderChunkDiagnostics {
            SourceKind = "rtf",
            TableCount = block.Tables?.Count ?? 0,
            ImageCount = imageCount,
            SelectedPageCount = 0,
            PageCount = 0,
            LinkCount = CountHyperlinkRuns(document),
            FormFieldCount = CountFormFields(document)
        };
    }

    private static int CountHyperlinkRuns(RtfDocument document) {
        int count = 0;
        for (int i = 0; i < document.Paragraphs.Count; i++) {
            RtfParagraph paragraph = document.Paragraphs[i];
            for (int runIndex = 0; runIndex < paragraph.Runs.Count; runIndex++) {
                if (paragraph.Runs[runIndex].Hyperlink != null) count++;
            }
        }

        return count;
    }

    private static int CountFormFields(RtfDocument document) {
        int count = 0;
        for (int i = 0; i < document.Paragraphs.Count; i++) {
            RtfParagraph paragraph = document.Paragraphs[i];
            for (int inlineIndex = 0; inlineIndex < paragraph.Inlines.Count; inlineIndex++) {
                if (paragraph.Inlines[inlineIndex] is RtfField field && field.FormFieldData != null) {
                    count++;
                }
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
