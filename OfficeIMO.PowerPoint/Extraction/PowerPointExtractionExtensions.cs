using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace OfficeIMO.PowerPoint;

/// <summary>
/// Chunked extraction helpers intended for AI ingestion.
/// </summary>
public static class PowerPointExtractionExtensions {
    /// <summary>
    /// Options controlling PowerPoint extraction behavior.
    /// </summary>
    public sealed class PowerPointExtractOptions {
        /// <summary>
        /// When true, include speaker notes in output. Default: true.
        /// </summary>
        public bool IncludeNotes { get; set; } = true;

        /// <summary>
        /// When true, include slide tables in output. Default: true.
        /// </summary>
        public bool IncludeTables { get; set; } = true;

        /// <summary>
        /// When true, include hidden shapes in output. Default: true.
        /// </summary>
        public bool IncludeHiddenShapes { get; set; } = true;
    }

    /// <summary>
    /// Extracts a presentation into slide-aligned chunks (one chunk per slide by default).
    /// </summary>
    public static IEnumerable<PowerPointExtractChunk> ExtractMarkdownChunks(
        this PowerPointPresentation presentation,
        PowerPointExtractOptions? extract = null,
        PowerPointExtractChunkingOptions? chunking = null,
        string? sourcePath = null,
        CancellationToken cancellationToken = default) {
        if (presentation == null) throw new ArgumentNullException(nameof(presentation));
        extract ??= new PowerPointExtractOptions();
        chunking ??= new PowerPointExtractChunkingOptions();
        if (chunking.MaxChars < 256) chunking.MaxChars = 256;

        for (int i = 0; i < presentation.Slides.Count; i++) {
            cancellationToken.ThrowIfCancellationRequested();

            var slide = presentation.Slides[i];
            int slideNumber = i + 1;
            List<string>? warnings = null;

            var md = new StringBuilder();
            md.Append("## Slide ").AppendLine(slideNumber.ToString(System.Globalization.CultureInfo.InvariantCulture));
            md.AppendLine();

            // TextBoxes in shape order.
            foreach (var tb in slide.TextBoxes) {
                cancellationToken.ThrowIfCancellationRequested();
                if (!extract.IncludeHiddenShapes && tb.Hidden) continue;
                var text = (tb.Text ?? string.Empty).Trim();
                if (text.Length == 0) continue;
                md.AppendLine(text);
                md.AppendLine();
            }

            if (extract.IncludeTables) {
                AppendTablesMarkdown(md, slide, extract);
            }

            if (extract.IncludeNotes) {
                if (slide.Notes.TryGetExistingText(out string notes)) {
                    md.AppendLine("### Notes");
                    md.AppendLine();
                    md.AppendLine(notes);
                    md.AppendLine();
                }
            }

            var markdown = md.ToString().TrimEnd();
            if (markdown.Length > chunking.MaxChars) {
                markdown = markdown.Substring(0, chunking.MaxChars) + "\n\n<!-- truncated -->";
                warnings = new List<string> { "Markdown truncated to MaxChars." };
            }

            var id = BuildStableId("ppt-md", sourcePath, slideNumber);
            yield return new PowerPointExtractChunk {
                Id = id,
                Location = new PowerPointExtractLocation {
                    Path = sourcePath,
                    Slide = slideNumber,
                    BlockIndex = i
                },
                Text = markdown,
                Markdown = markdown,
                Warnings = warnings
            };
        }
    }

    private static void AppendTablesMarkdown(StringBuilder markdown, PowerPointSlide slide, PowerPointExtractOptions extract) {
        int tableIndex = 0;
        foreach (var table in slide.Tables) {
            if (!extract.IncludeHiddenShapes && table.Hidden) continue;
            tableIndex++;
            List<IReadOnlyList<string>> rows = table.RowItems
                .Select(row => (IReadOnlyList<string>)row.Cells.Select(cell => NormalizeText(cell.Text)).ToList())
                .ToList();

            if (rows.Count == 0 || rows.All(row => row.All(string.IsNullOrWhiteSpace))) {
                continue;
            }

            int columnCount = rows.Max(row => row.Count);
            if (columnCount == 0) {
                continue;
            }

            markdown.Append("### Table ").AppendLine(tableIndex.ToString(System.Globalization.CultureInfo.InvariantCulture));
            markdown.AppendLine();
            AppendMarkdownTable(markdown, rows, columnCount);
            markdown.AppendLine();
        }
    }

    private static void AppendMarkdownTable(StringBuilder markdown, IReadOnlyList<IReadOnlyList<string>> rows, int columnCount) {
        IReadOnlyList<string> header = rows[0];
        markdown.Append('|');
        for (int i = 0; i < columnCount; i++) {
            markdown.Append(' ').Append(EscapeMarkdownTableCell(GetCellValue(header, i, $"Column {i + 1}"))).Append(" |");
        }
        markdown.AppendLine();

        markdown.Append('|');
        for (int i = 0; i < columnCount; i++) {
            markdown.Append(" --- |");
        }
        markdown.AppendLine();

        foreach (IReadOnlyList<string> row in rows.Skip(1)) {
            markdown.Append('|');
            for (int i = 0; i < columnCount; i++) {
                markdown.Append(' ').Append(EscapeMarkdownTableCell(GetCellValue(row, i, string.Empty))).Append(" |");
            }
            markdown.AppendLine();
        }
    }

    private static string GetCellValue(IReadOnlyList<string> row, int index, string fallback) {
        if (index >= row.Count) {
            return fallback;
        }

        string value = row[index];
        return string.IsNullOrWhiteSpace(value) ? fallback : value;
    }

    private static string EscapeMarkdownTableCell(string value) {
        return (value ?? string.Empty)
            .Replace("\r\n", " ")
            .Replace('\n', ' ')
            .Replace('\r', ' ')
            .Replace("|", "\\|");
    }

    private static string NormalizeText(string? text) {
        return (text ?? string.Empty).Trim();
    }

    private static string BuildStableId(string kind, string? path, int slideNumber) {
        var safe = string.IsNullOrWhiteSpace(path) ? "memory" : System.IO.Path.GetFileName(path!.Trim());
        return $"{kind}:{safe}:s{slideNumber}";
    }
}
