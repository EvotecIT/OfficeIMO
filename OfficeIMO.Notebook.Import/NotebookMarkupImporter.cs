using System.Globalization;

namespace OfficeIMO.Notebook.Import;

internal static partial class NotebookMarkupImporter {
    private const int MaxAuthoringMetadataCharacters = 1024 * 1024;
    internal static void ReadQuarto(ReadOnlyMemory<byte> input, NotebookImportContext context) =>
        ReadSource(NotebookImportContext.DecodeUtf8(input), NotebookImportFormat.Quarto, context);

    internal static void ReadMyST(ReadOnlyMemory<byte> input, NotebookImportContext context) =>
        ReadSource(NotebookImportContext.DecodeUtf8(input), NotebookImportFormat.MySTMarkdown, context);

    internal static void ReadRMarkdown(ReadOnlyMemory<byte> input, NotebookImportContext context) =>
        ReadSource(NotebookImportContext.DecodeUtf8(input), NotebookImportFormat.RMarkdown, context);

    internal static void ReadSource(string source, NotebookImportFormat format, NotebookImportContext context) {
        string? language = format is NotebookImportFormat.RMarkdown or NotebookImportFormat.RNotebookHtml ? "r" : null;
        var body = new SourceCursor(source, source.Length > 0 && source[0] == '\uFEFF' ? 1 : 0);
        ReadFrontMatter(ref body, ref language, context);
        if (language != null) SetDocumentLanguage(context.Document, language);
        int proseStart = body.Offset;
        bool insideHtmlComment = false;
        while (body.TryRead(out SourceLine line)) {
            context.CancellationToken.ThrowIfCancellationRequested();
            ReadOnlySpan<char> text = line.Text(source);
            if (IsHtmlCommentBlock(text, ref insideHtmlComment)) continue;
            if (!TryOpenFence(text, out Fence fence)) {
                ReportInlineSyntax(text, format, line, context);
                continue;
            }

            ReadOnlySpan<char> info = text.Slice(fence.InfoOffset).Trim();
            bool executable = TryGetLanguage(info, format, language, out string? cellLanguage);
            int codeStart = body.Offset;
            if (!TryFindClose(ref body, fence, context.CancellationToken, out int codeEnd)) {
                if (executable) {
                    context.Report("notebook-import-open-fence", "The executable fence has no closing delimiter. Its content was retained as Markdown.", Location(line.Number));
                }
                break;
            }
            if (!executable) {
                if (IsBracedDirective(info, format)) {
                    context.Report("notebook-import-markup-directive", "A document-specific fenced directive was retained as source.", Location(line.Number),
                        OfficeConversionLossKind.Approximation);
                }
                continue;
            }

            AddProse(source, proseStart, line.Start, context);
            int executableStart = ReadCellOptions(source, codeStart, codeEnd, format, cellLanguage, context.CancellationToken);
            if (info.Length > MaxAuthoringMetadataCharacters) {
                context.Report("notebook-import-chunk-metadata-limit", "An oversized chunk header was retained in a raw cell.", Location(line.Number),
                    OfficeConversionLossKind.Approximation);
                context.AddCell(NotebookCellKind.Raw, info.ToString());
            }
            if (executableStart - codeStart > MaxAuthoringMetadataCharacters) {
                context.Report("notebook-import-chunk-metadata-limit", "Oversized chunk options were retained with the code source.", Location(line.Number),
                    OfficeConversionLossKind.Approximation);
                executableStart = codeStart;
            }
            NotebookCell cell = context.AddCell(NotebookCellKind.Code,
                ReadFenceContent(source, executableStart, codeEnd, fence.Indent, context.CancellationToken), cellLanguage);
            if (info.Length <= MaxAuthoringMetadataCharacters) {
                cell.SetMetadata("notebook_import_header", NotebookImportContext.JsonString(info.ToString()));
            }
            if (executableStart > codeStart) {
                cell.SetMetadata("notebook_import_options", NotebookImportContext.JsonString(
                    ReadFenceContent(source, codeStart, executableStart, fence.Indent, context.CancellationToken)));
            }
            bool hasHeaderOptions = (format is NotebookImportFormat.RMarkdown or NotebookImportFormat.RNotebookHtml ||
                format == NotebookImportFormat.Quarto && string.Equals(cellLanguage, "r", StringComparison.OrdinalIgnoreCase))
                && info.IndexOf(',') >= 0;
            if (executableStart > codeStart || hasHeaderOptions) {
                context.Report("notebook-import-chunk-options",
                    "Chunk authoring options were retained as metadata. The import reads source and saved results without applying those options.",
                    Location(line.Number), OfficeConversionLossKind.Approximation);
            }
            proseStart = body.Offset;
        }
        AddProse(source, proseStart, source.Length, context);
    }

    private static void AddProse(string source, int start, int end, NotebookImportContext context) {
        if (source.AsSpan(start, end - start).Trim().IsEmpty) return;
        context.AddCell(NotebookCellKind.Markdown, source.Substring(start, end - start));
    }

    private static bool TryFindClose(ref SourceCursor body, Fence fence, CancellationToken token, out int codeEnd) {
        while (body.TryRead(out SourceLine line)) {
            token.ThrowIfCancellationRequested();
            if (IsClosingFence(line.Text(body.Source), fence)) {
                codeEnd = line.Start;
                return true;
            }
        }
        codeEnd = body.Offset;
        return false;
    }

    private static void ReportInlineSyntax(ReadOnlySpan<char> text, NotebookImportFormat format, SourceLine line,
        NotebookImportContext context) {
        ReadOnlySpan<char> trimmed = text.TrimStart();
        bool directive = trimmed.StartsWith(":::".AsSpan(), StringComparison.Ordinal)
            && trimmed.TrimStart(':').TrimStart().StartsWith("{".AsSpan(), StringComparison.Ordinal);
        bool inline = format switch {
            NotebookImportFormat.MySTMarkdown => text.IndexOf("{eval}`".AsSpan(), StringComparison.Ordinal) >= 0,
            NotebookImportFormat.RMarkdown or NotebookImportFormat.RNotebookHtml => text.IndexOf("`r ".AsSpan(), StringComparison.Ordinal) >= 0,
            NotebookImportFormat.Quarto => text.IndexOf("`{".AsSpan(), StringComparison.Ordinal) >= 0
                || text.IndexOf("{{<".AsSpan(), StringComparison.Ordinal) >= 0,
            _ => false
        };
        if (directive || inline) {
            context.Report("notebook-import-markup-directive", "A document directive or inline expression was retained as source without evaluation.",
                Location(line.Number), OfficeConversionLossKind.Approximation);
        }
    }

    private static string Location(int line) => $"lines[{line.ToString(CultureInfo.InvariantCulture)}]";
}
