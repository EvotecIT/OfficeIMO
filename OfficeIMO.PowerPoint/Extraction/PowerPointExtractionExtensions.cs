using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using System.Threading;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using A = DocumentFormat.OpenXml.Drawing;

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

            int tableIndex = 0;
            foreach (PowerPointShape shape in slide.EnumerateShapesDeep(
                         slide.Shapes, extract.IncludeHiddenShapes)) {
                cancellationToken.ThrowIfCancellationRequested();

                if (shape is PowerPointTextBox textBox) {
                    AppendTextBoxMarkdown(md, textBox);
                } else if (extract.IncludeTables && shape is PowerPointTable table) {
                    AppendTableMarkdown(md, table, ref tableIndex);
                }
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

    private static void AppendTextBoxMarkdown(StringBuilder markdown, PowerPointTextBox textBox) {
        bool isTitle = textBox.ShapePlaceholderType == PlaceholderValues.Title
            || textBox.ShapePlaceholderType == PlaceholderValues.CenteredTitle;
        var numberingState = new Dictionary<int, int>();
        var listContentIndents = new Dictionary<int, int>();
        bool appended = false;
        ParagraphMarkdownKind? previousKind = null;
        foreach (PowerPointParagraph paragraph in textBox.Paragraphs) {
            string text = RenderParagraphText(paragraph).Trim();
            if (text.Length == 0) continue;

            ParagraphMarkdownKind kind = isTitle
                ? ParagraphMarkdownKind.Heading
                : paragraph.IsNumbered
                    ? ParagraphMarkdownKind.NumberedList
                    : !string.IsNullOrEmpty(paragraph.BulletCharacter)
                        ? ParagraphMarkdownKind.BulletList
                        : ParagraphMarkdownKind.Plain;
            if (appended && previousKind.HasValue
                && (!IsListKind(previousKind.Value) || !IsListKind(kind))) {
                markdown.AppendLine();
            }

            if (kind == ParagraphMarkdownKind.Heading) {
                markdown.Append("### ").AppendLine(text);
            } else {
                int level = Math.Max(0, paragraph.Level ?? 0);
                int indent = ResolveListIndent(level, listContentIndents);
                if (kind == ParagraphMarkdownKind.NumberedList) {
                    int number = ResolveNumberingValue(paragraph, level, numberingState);
                    string markdownMarker = number.ToString(CultureInfo.InvariantCulture) + ".";
                    string powerPointMarker = PowerPointNumberingFormatter.FormatMarker(
                        number, paragraph.NumberingScheme);
                    markdown.Append(' ', indent);
                    markdown.Append(markdownMarker).Append(' ');
                    // CommonMark only recognizes decimal ordered-list markers.
                    // Retain the displayed PowerPoint marker as visible text.
                    if (!string.Equals(markdownMarker, powerPointMarker, StringComparison.Ordinal)) {
                        markdown.Append(powerPointMarker).Append(' ');
                    }
                    markdown.AppendLine(text);
                    UpdateListContentIndent(listContentIndents, level,
                        checked(indent + markdownMarker.Length + 1));
                } else if (kind == ParagraphMarkdownKind.BulletList) {
                    char markdownMarker = ResolveBulletMarker(
                        paragraph.BulletCharacter);
                    markdown.Append(' ', indent);
                    markdown.Append(markdownMarker).Append(' ');
                    if (!string.Equals(markdownMarker.ToString(),
                            paragraph.BulletCharacter,
                            StringComparison.Ordinal)) {
                        markdown.Append(paragraph.BulletCharacter)
                            .Append(' ');
                    }
                    markdown.AppendLine(text);
                    UpdateListContentIndent(listContentIndents, level,
                        checked(indent + 2));
                } else {
                    markdown.AppendLine(text);
                }
            }
            appended = true;
            previousKind = kind;
        }

        if (appended) {
            markdown.AppendLine();
        }
    }

    private static int ResolveNumberingValue(PowerPointParagraph paragraph, int level,
        IDictionary<int, int> numberingState) {
        int number = paragraph.NumberingStartAt
            ?? (numberingState.TryGetValue(level, out int previous) ? previous + 1 : 1);
        numberingState[level] = number;
        return number;
    }

    private static bool IsListKind(ParagraphMarkdownKind kind) =>
        kind == ParagraphMarkdownKind.BulletList || kind == ParagraphMarkdownKind.NumberedList;

    private static string RenderParagraphText(PowerPointParagraph paragraph) {
        IReadOnlyList<PowerPointTextRun> runs = paragraph.Runs;
        int runIndex = 0;
        var text = new StringBuilder();
        foreach (OpenXmlElement child in paragraph.Paragraph.ChildElements) {
            if (child is A.Run) {
                PowerPointTextRun? run = runIndex < runs.Count ? runs[runIndex] : null;
                runIndex++;
                string runText = run?.Text ?? child.InnerText ?? string.Empty;
                if (run?.Hyperlink != null && runText.Length > 0) {
                    text.Append('[').Append(EscapeMarkdownLinkLabel(runText)).Append("](<")
                        .Append(EscapeMarkdownLinkDestination(run.Hyperlink.ToString())).Append(">)");
                } else {
                    text.Append(runText);
                }
            } else if (child is A.Break) {
                text.Append("<br />");
            } else if (child is A.Field field) {
                text.Append(field.Text?.Text ?? field.InnerText ?? string.Empty);
            }
        }

        return text.Length == 0 ? paragraph.Text : text.ToString();
    }

    private static char ResolveBulletMarker(string? bulletCharacter) {
        return bulletCharacter is "*" or "+" or "-" ? bulletCharacter[0] : '-';
    }

    private static int ResolveListIndent(int level,
        IReadOnlyDictionary<int, int> listContentIndents) {
        if (level <= 0) return 0;
        int conventionalIndent = checked(level * 4);
        return listContentIndents.TryGetValue(level - 1,
                out int parentContentIndent)
            ? Math.Max(conventionalIndent, parentContentIndent)
            : conventionalIndent;
    }

    private static void UpdateListContentIndent(
        IDictionary<int, int> listContentIndents, int level,
        int contentIndent) {
        foreach (int deeperLevel in listContentIndents.Keys
                     .Where(existingLevel => existingLevel > level)
                     .ToArray()) {
            listContentIndents.Remove(deeperLevel);
        }
        listContentIndents[level] = contentIndent;
    }

    private static string EscapeMarkdownLinkLabel(string value) {
        return value.Replace("\\", "\\\\")
            .Replace("[", "\\[")
            .Replace("]", "\\]");
    }

    private static string EscapeMarkdownLinkDestination(string value) {
        var escaped = new StringBuilder(value.Length);
        foreach (char character in value) {
            if (char.IsWhiteSpace(character) || char.IsControl(character)
                || character is '<' or '>' or '\\') {
                escaped.Append(Uri.EscapeDataString(character.ToString()));
            } else {
                escaped.Append(character);
            }
        }
        return escaped.ToString();
    }

    private static void AppendTableMarkdown(StringBuilder markdown, PowerPointTable table, ref int tableIndex) {
        List<IReadOnlyList<string>> rows = table.RowItems
            .Select(row => (IReadOnlyList<string>)row.Cells.Select(RenderTableCellText).ToList())
            .ToList();

        if (rows.Count == 0 || rows.All(row => row.All(string.IsNullOrWhiteSpace))) {
            return;
        }

        int columnCount = rows.Max(row => row.Count);
        if (columnCount == 0) {
            return;
        }

        tableIndex++;
        markdown.Append("### Table ").AppendLine(tableIndex.ToString(CultureInfo.InvariantCulture));
        markdown.AppendLine();
        AppendMarkdownTable(markdown, rows, columnCount);
        markdown.AppendLine();
    }

    private static string RenderTableCellText(PowerPointTableCell cell) {
        string[] paragraphs = cell.Paragraphs
            .Select(paragraph => RenderParagraphText(paragraph).Trim())
            .Where(text => text.Length > 0)
            .ToArray();
        return paragraphs.Length == 0
            ? NormalizeText(cell.Text)
            : string.Join("<br />", paragraphs);
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

    private enum ParagraphMarkdownKind {
        Plain,
        Heading,
        BulletList,
        NumberedList
    }
}
