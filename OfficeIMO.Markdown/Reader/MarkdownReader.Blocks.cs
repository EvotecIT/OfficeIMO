using System.Globalization;
using System.Linq;

namespace OfficeIMO.Markdown;

/// <summary>
/// Block parsing helpers for <see cref="MarkdownReader"/>.
/// </summary>
public static partial class MarkdownReader {
    private static bool IsAtxHeading(string line, out int level, out string text) {
        level = 0; text = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        int i = 0; while (i < line.Length && line[i] == '#') i++;
        if (i < 1 || i > 6) return false;
        if (i < line.Length && line[i] == ' ') { level = i; text = line.Substring(i + 1); return true; }
        return false;
    }

    private static bool IsCodeFenceOpen(string line, out string language, out char fenceChar, out int fenceLength) {
        language = string.Empty; fenceChar = '\0'; fenceLength = 0;
        if (line is null) return false;
        line = line.Trim();
        if (line.Length < 3) return false;
        char ch = line[0];
        if (ch != '`' && ch != '~') return false;

        int run = 0;
        while (run < line.Length && line[run] == ch) run++;
        if (run < 3) return false;

        fenceChar = ch;
        fenceLength = run;
        language = line.Length > run ? line.Substring(run).Trim() : string.Empty;
        return true;
    }
    private static bool IsCodeFenceClose(string line, char fenceChar, int fenceLength) {
        if (line is null) return false;
        var trimmed = line.Trim();
        if (trimmed.Length < Math.Max(3, fenceLength)) return false;
        // CommonMark allows closing fence length >= opening fence length. We accept that.
        for (int i = 0; i < trimmed.Length; i++) {
            if (trimmed[i] != fenceChar) return false;
        }
        return trimmed.Length >= Math.Max(3, fenceLength);
    }

    private static bool TryParseCaption(string line, out string caption) {
        caption = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.Trim();
        if (t.Length >= 3 && t[0] == '_' && t[t.Length - 1] == '_' && t.IndexOf('_', 1) == t.Length - 1) { caption = t.Substring(1, t.Length - 2); return true; }
        return false;
    }

    private static bool IsImageLine(string line) => TryParseImage(line, out _, out _);
    private static bool TryParseImage(string line, out ImageBlock image, out string? sizeSpec) {
        image = null!;
        sizeSpec = null;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.Trim();
        if (!t.StartsWith("![")) return false;
        int altEnd = t.IndexOf(']');
        if (altEnd < 2) return false;
        if (altEnd + 1 >= t.Length || t[altEnd + 1] != '(') return false;
        int parenClose = t.IndexOf(')', altEnd + 2);
        if (parenClose <= altEnd + 2) return false;
        string alt = t.Substring(2, altEnd - 2);
        string inner = t.Substring(altEnd + 2, parenClose - (altEnd + 2));
        if (!TrySplitUrlAndOptionalTitle(inner, out var src, out var title)) {
            src = inner.Trim();
            title = null;
        }
        image = new ImageBlock(src, alt, title);
        // Optional attribute list: {width=.. height=..}
        if (parenClose + 1 < t.Length) {
            var rest = t.Substring(parenClose + 1).Trim();
            if (rest.StartsWith("{")) {
                int close = rest.IndexOf('}');
                if (close > 0) {
                    sizeSpec = rest.Substring(1, close - 1).Trim();
                    var attrs = sizeSpec;
                    foreach (var part in attrs.Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)) {
                        int eq = part.IndexOf('=');
                        if (eq > 0) {
                            var key = part.Substring(0, eq).Trim();
                            var val = part.Substring(eq + 1).Trim();
                            if (double.TryParse(val, System.Globalization.NumberStyles.Number, System.Globalization.CultureInfo.InvariantCulture, out var num)) {
                                if (string.Equals(key, "width", StringComparison.OrdinalIgnoreCase)) image.Width = num;
                                else if (string.Equals(key, "height", StringComparison.OrdinalIgnoreCase)) image.Height = num;
                            }
                        }
                    }
                }
            }
        }
        return true;
    }

    /// <summary>
    /// Determines whether a line is likely to be part of a markdown table. The logic follows
    /// CommonMark's relaxed table rules: when both outer pipes are present a single column is
    /// permitted, otherwise at least two pipe-separated cells are required so that plain
    /// paragraphs containing a single <c>|</c> are not mis-classified as tables.
    /// </summary>
    private static bool LooksLikeTableRow(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        var trimmed = line.Trim();
        if (trimmed.Length < 3 || !trimmed.Contains('|')) return false;

        var cells = SplitTableRow(trimmed);
        if (cells.Count == 0) return false;

        bool hasLeadingPipe = trimmed[0] == '|';
        bool hasTrailingPipe = trimmed[trimmed.Length - 1] == '|';
        bool hasOuterPipes = hasLeadingPipe && hasTrailingPipe;

        if (!hasOuterPipes && cells.Count < 2) return false;

        return true;
    }

    private static bool StartsTable(string[] lines, int index) => TryGetTableExtent(lines, index, out _, out _);

    private static bool TryGetTableExtent(string[] lines, int start, out int end, out bool hasOuterPipes) {
        end = start;
        hasOuterPipes = false;
        if (lines is null || start < 0 || start >= lines.Length) return false;
        if (!LooksLikeTableRow(lines[start])) return false;
        if (IsAlignmentRow(lines[start])) return false;

        var firstTrimmed = lines[start].Trim();
        if (firstTrimmed.Length == 0) return false;

        bool hasLeadingPipe = firstTrimmed[0] == '|';
        bool hasTrailingPipe = firstTrimmed[firstTrimmed.Length - 1] == '|';
        hasOuterPipes = hasLeadingPipe && hasTrailingPipe;

        int j = start + 1;
        bool sawAlignmentRow = false;
        if (j < lines.Length && IsAlignmentRow(lines[j])) {
            sawAlignmentRow = true;
            j++;
        }

        if (!sawAlignmentRow) {
            // Headerless tables are easy to mis-detect (any two lines with pipes). To reduce false positives,
            // require explicit outer pipes on every row and at least 2 rows.
            if (!hasOuterPipes) return false;

            // Require the 2nd row to also have outer pipes, otherwise treat the first row as a paragraph line.
            if (j >= lines.Length) return false;
            var second = (lines[j] ?? string.Empty).Trim();
            if (!(second.Length > 0 && second[0] == '|' && second[second.Length - 1] == '|')) return false;

            while (j < lines.Length) {
                var t = (lines[j] ?? string.Empty).Trim();
                if (t.Length == 0) break;
                if (!LooksLikeTableRow(t)) break;
                if (!(t[0] == '|' && t[t.Length - 1] == '|')) break;
                j++;
            }
        } else {
            while (j < lines.Length && LooksLikeTableRow(lines[j])) j++;
        }

        end = j - 1;
        return true;
    }

    private static TableBlock ParseTable(string[] lines, int start, int end) {
        var cells0 = SplitTableRow(lines[start]);
        var table = new TableBlock();
        if (start + 1 <= end && IsAlignmentRow(lines[start + 1])) {
            table.Headers.AddRange(cells0);
            var aligns = SplitTableRow(lines[start + 1]);
            for (int i = 0; i < aligns.Count; i++) table.Alignments.Add(ParseAlignmentCell(aligns[i]));
            for (int i = start + 2; i <= end; i++) table.Rows.Add(SplitTableRow(lines[i]));
        } else {
            for (int i = start; i <= end; i++) table.Rows.Add(SplitTableRow(lines[i]));
        }
        return table;
    }

    private static bool IsAlignmentRow(string line) {
        var cells = SplitTableRow(line);
        if (cells.Count == 0) return false;
        foreach (var c in cells) {
            var t = c.Trim(); if (t.Length < 3) return false;
            int dash = 0;
            for (int i = 0; i < t.Length; i++) {
                char ch = t[i];
                if (ch == '-') dash++;
                else if (ch == ':' && (i == 0 || i == t.Length - 1)) { } else return false;
            }
            if (dash < 3) return false;
        }
        return true;
    }

    private static ColumnAlignment ParseAlignmentCell(string cell) {
        var t = cell.Trim();
        if (t.StartsWith(":")) { if (t.EndsWith(":")) return ColumnAlignment.Center; return ColumnAlignment.Left; }
        if (t.EndsWith(":")) return ColumnAlignment.Right;
        return ColumnAlignment.None;
    }

    private static List<string> SplitTableRow(string line) {
        if (line is null) return new List<string>();
        var t = line.Trim();
        if (t.StartsWith("|")) t = t.Substring(1);
        if (t.EndsWith("|") && t.Length > 0) t = t.Substring(0, t.Length - 1);

        // Split on unescaped pipes that are not inside backtick code spans.
        // This covers the most common GFM-table cases:
        // - Escaped pipes: \|
        // - Pipes inside code spans: `a|b` or ``a|b``
        var cells = new List<string>();
        var sb = new StringBuilder(t.Length);
        int i = 0;
        int codeFenceLen = 0; // 0 = not in code span

        while (i < t.Length) {
            char ch = t[i];

            // Backslash escape: keep as-is, but prevent the next char (including '|') from being treated specially.
            if (ch == '\\' && i + 1 < t.Length) {
                sb.Append(ch);
                sb.Append(t[i + 1]);
                i += 2;
                continue;
            }

            // Code span tracking: toggle on/off when encountering a run of backticks of the same length.
            if (ch == '`') {
                int run = 1;
                int j = i + 1;
                while (j < t.Length && t[j] == '`') { run++; j++; }

                if (codeFenceLen == 0) codeFenceLen = run;
                else if (run == codeFenceLen) codeFenceLen = 0;

                sb.Append(t, i, run);
                i += run;
                continue;
            }

            if (ch == '|' && codeFenceLen == 0) {
                cells.Add(sb.ToString().Trim());
                sb.Clear();
                i++;
                continue;
            }

            sb.Append(ch);
            i++;
        }

        cells.Add(sb.ToString().Trim());

        // Unescape backslash escapes outside code spans for cell storage, so that doc->ToMarkdown roundtrips cleanly.
        for (int c = 0; c < cells.Count; c++) {
            cells[c] = UnescapeBackslashEscapesOutsideCodeSpans(cells[c]);
        }

        return cells;
    }

    private static string UnescapeBackslashEscapesOutsideCodeSpans(string value) {
        if (string.IsNullOrEmpty(value)) return value ?? string.Empty;

        var sb = new StringBuilder(value.Length);
        int i = 0;
        int codeFenceLen = 0;

        while (i < value.Length) {
            char ch = value[i];

            if (ch == '`') {
                int run = 1;
                int j = i + 1;
                while (j < value.Length && value[j] == '`') { run++; j++; }

                if (codeFenceLen == 0) codeFenceLen = run;
                else if (run == codeFenceLen) codeFenceLen = 0;

                sb.Append(value, i, run);
                i += run;
                continue;
            }

            if (ch == '\\' && codeFenceLen == 0 && i + 1 < value.Length) {
                char next = value[i + 1];
                if (IsBackslashEscapable(next)) {
                    sb.Append(next);
                    i += 2;
                    continue;
                }
            }

            sb.Append(ch);
            i++;
        }

        return sb.ToString();
    }

    private static int CountLeadingSpaces(string line) {
        if (string.IsNullOrEmpty(line)) return 0;
        int i = 0;
        while (i < line.Length && line[i] == ' ') i++;
        return i;
    }

    private static List<string> ConsumeListContinuationLines(string[] lines, ref int nextIndex, int itemLevel, string initialContent, MarkdownReaderOptions options) {
        if (lines == null) return new List<string> { initialContent ?? string.Empty };
        if (nextIndex < 0) nextIndex = 0;

        // Heuristic: list item content begins after marker; our parsers model indentation levels as 2 spaces per level.
        int baseIndent = Math.Max(0, itemLevel) * 2;
        int continuationIndent = baseIndent + 2;

        var collected = new List<string> { initialContent ?? string.Empty };
        int k = nextIndex;

        while (k < lines.Length) {
            var line = lines[k] ?? string.Empty;

            // Stop before the next list item (including nested items).
            if (IsUnorderedListLine(line, out _, out _, out _, out _) ||
                IsOrderedListLine(line, out _, out _, out _)) {
                break;
            }

            // Stop before nested blocks; they are handled as child blocks of the list item.
            if (continuationIndent <= line.Length) {
                var slice = line.Substring(Math.Min(continuationIndent, line.Length));
                var sliceTrim = slice.TrimStart();
                if (IsCodeFenceOpen(slice, out _, out _, out _)) break;
                if (sliceTrim.StartsWith(">")) break;

                if (options.HtmlBlocks && sliceTrim.StartsWith("<")) {
                    // Avoid breaking on angle-bracket autolinks like "<https://...>".
                    if (!TryParseAngleAutolink(sliceTrim, 0, out _, out _, out _)) break;
                }

                // Indented code block inside list item: continuationIndent + 4 spaces.
                if (options.IndentedCodeBlocks) {
                    int spacesIndent = CountLeadingSpaces(line);
                    if (spacesIndent >= continuationIndent + 4) break;
                }

                // Table inside list item: a pipe row followed by an alignment/row.
                if (options.Tables && LooksLikeTableRow(sliceTrim)) {
                    int peek = k + 1;
                    if (peek < lines.Length && CountLeadingSpaces(lines[peek] ?? string.Empty) >= continuationIndent) {
                        var nextSlice = (lines[peek] ?? string.Empty).Substring(Math.Min(continuationIndent, (lines[peek] ?? string.Empty).Length)).TrimStart();
                        // Reduce false positives: require an alignment row, or explicit outer pipes on both rows.
                        bool curOuter = sliceTrim.Length > 0 && sliceTrim[0] == '|' && sliceTrim[sliceTrim.Length - 1] == '|';
                        bool nextOuter = nextSlice.Length > 0 && nextSlice[0] == '|' && nextSlice[nextSlice.Length - 1] == '|';
                        if (IsAlignmentRow(nextSlice) || (curOuter && nextOuter)) break;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(line)) {
                // Keep blank lines only if followed by an indented continuation line; otherwise end item.
                int peek = k + 1;
                if (peek >= lines.Length) break;
                var next = lines[peek] ?? string.Empty;
                if (IsUnorderedListLine(next, out _, out _, out _, out _) ||
                    IsOrderedListLine(next, out _, out _, out _)) {
                    break;
                }
                int nextSpaces = CountLeadingSpaces(next);
                if (nextSpaces < continuationIndent) break;

                collected.Add(string.Empty);
                k++;
                continue;
            }

            int spaces = CountLeadingSpaces(line);
            if (spaces < continuationIndent) break;

            // Strip the required indent; keep the remainder as-is (including additional indentation).
            string cont = line.Length >= continuationIndent ? line.Substring(continuationIndent) : string.Empty;
            cont = cont.TrimStart();
            collected.Add(cont);
            k++;
        }

        nextIndex = k;
        return collected;
    }

    private static bool TryParseNestedFencedCodeBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, out CodeBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.FencedCode) return false;

        string line = lines[index] ?? string.Empty;
        int indent = CountLeadingSpaces(line);
        if (indent < continuationIndent) return false;

        string first = line.Length >= continuationIndent ? line.Substring(continuationIndent) : string.Empty;
        if (!IsCodeFenceOpen(first, out string language, out char fenceChar, out int fenceLen)) return false;

        int j = index + 1;
        var code = new StringBuilder();
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            int ind = CountLeadingSpaces(raw);
            string sliced = ind >= continuationIndent && raw.Length >= continuationIndent ? raw.Substring(continuationIndent) : raw.TrimStart();
            if (IsCodeFenceClose(sliced, fenceChar, fenceLen)) { j++; break; }
            code.AppendLine(sliced);
            j++;
        }

        var cb = new CodeBlock(language, code.ToString().TrimEnd('\r', '\n'));
        // Optional caption line (indented like other nested content)
        if (j < lines.Length) {
            var capLine = lines[j] ?? string.Empty;
            if (CountLeadingSpaces(capLine) >= continuationIndent) {
                var capSlice = capLine.Length >= continuationIndent ? capLine.Substring(continuationIndent) : capLine.TrimStart();
                if (TryParseCaption(capSlice, out var cap)) { cb.Caption = cap; j++; }
            }
        }

        block = cb;
        index = j;
        return true;
    }

    private static bool TryParseNestedIndentedCodeBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, out CodeBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.IndentedCodeBlocks) return false;

        string line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) return false;

        int spaces = CountLeadingSpaces(line);
        int required = continuationIndent + 4;
        if (spaces < required) return false;

        int j = index;
        var sb = new StringBuilder();
        while (j < lines.Length) {
            string cur = lines[j] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cur)) {
                int peek = j + 1;
                if (peek >= lines.Length) break;
                int nextSpaces = CountLeadingSpaces(lines[peek] ?? string.Empty);
                if (nextSpaces < required) break;
                sb.AppendLine();
                j++;
                continue;
            }

            int curSpaces = CountLeadingSpaces(cur);
            if (curSpaces < required) break;
            sb.AppendLine(cur.Substring(required));
            j++;
        }

        block = new CodeBlock(string.Empty, sb.ToString().TrimEnd('\r', '\n'));
        index = j;
        return true;
    }

    private static bool TryParseNestedQuoteBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, MarkdownReaderState state, out QuoteBlock? quote) {
        quote = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;

        string line = lines[index] ?? string.Empty;
        if (CountLeadingSpaces(line) < continuationIndent) return false;
        string slice = line.Length >= continuationIndent ? line.Substring(continuationIndent) : string.Empty;
        if (!slice.TrimStart().StartsWith(">")) return false;

        int j = index;
        var collected = new List<string>();
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw)) {
                int peek = j + 1;
                if (peek >= lines.Length) break;
                var next = lines[peek] ?? string.Empty;
                if (CountLeadingSpaces(next) < continuationIndent) break;
                string nextPart = next.Length >= continuationIndent ? next.Substring(continuationIndent) : string.Empty;
                if (!nextPart.TrimStart().StartsWith(">")) break;
                collected.Add(string.Empty);
                j++;
                continue;
            }

            if (CountLeadingSpaces(raw) < continuationIndent) break;
            string part = raw.Length >= continuationIndent ? raw.Substring(continuationIndent) : string.Empty;

            if (string.IsNullOrWhiteSpace(part)) {
                int peek = j + 1;
                if (peek >= lines.Length) break;
                var next = lines[peek] ?? string.Empty;
                if (CountLeadingSpaces(next) < continuationIndent) break;
                string nextPart = next.Length >= continuationIndent ? next.Substring(continuationIndent) : string.Empty;
                if (!nextPart.TrimStart().StartsWith(">")) break;
                collected.Add(string.Empty);
                j++;
                continue;
            }

            if (!part.TrimStart().StartsWith(">")) break;
            collected.Add(part);
            j++;
        }

        if (collected.Count == 0) return false;

        var nested = ParseBlocksFromLines(collected.ToArray(), options, state);
        if (nested.Count > 0 && nested[0] is QuoteBlock qb) {
            quote = qb;
            index = j;
            return true;
        }
        return false;
    }

    private static bool TryParseNestedTableBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, MarkdownReaderState state, out TableBlock? table) {
        table = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.Tables) return false;

        string line = lines[index] ?? string.Empty;
        if (CountLeadingSpaces(line) < continuationIndent) return false;
        string slice = line.Length >= continuationIndent ? line.Substring(continuationIndent) : string.Empty;
        if (!LooksLikeTableRow(slice.TrimStart())) return false;

        int j = index;
        var collected = new List<string>();
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            if (CountLeadingSpaces(raw) < continuationIndent) break;
            string part = raw.Length >= continuationIndent ? raw.Substring(continuationIndent) : string.Empty;
            if (string.IsNullOrWhiteSpace(part)) break;
            // Stop when the row no longer looks table-ish.
            if (!LooksLikeTableRow(part.TrimStart()) && !IsAlignmentRow(part.TrimStart())) break;
            collected.Add(part);
            j++;
        }

        if (collected.Count == 0) return false;
        var nested = ParseBlocksFromLines(collected.ToArray(), options, state);
        if (nested.Count > 0 && nested[0] is TableBlock tb) {
            table = tb;
            index = j;
            return true;
        }
        return false;
    }

    private static bool TryParseNestedHtmlBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, MarkdownReaderState state, out IMarkdownBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.HtmlBlocks) return false;

        string line = lines[index] ?? string.Empty;
        if (CountLeadingSpaces(line) < continuationIndent) return false;
        string slice = line.Length >= continuationIndent ? line.Substring(continuationIndent) : string.Empty;
        string sliceTrim = slice.TrimStart();
        if (!sliceTrim.StartsWith("<")) return false;
        if (TryParseAngleAutolink(sliceTrim, 0, out _, out _, out _)) return false;

        // Collect contiguous indented lines and let HtmlBlockParser decide the extent.
        int j = index;
        var collected = new List<string>();
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw)) {
                // Allow unindented blank lines inside HTML blocks within list items.
                collected.Add(string.Empty);
                j++;
                continue;
            }
            if (CountLeadingSpaces(raw) < continuationIndent) break;
            collected.Add(raw.Length >= continuationIndent ? raw.Substring(continuationIndent) : raw.TrimStart());
            j++;
        }
        if (collected.Count == 0) return false;

        int local = 0;
        var tempDoc = MarkdownDoc.Create();
        var parser = new HtmlBlockParser();
        if (!parser.TryParse(collected.ToArray(), ref local, options, tempDoc, state)) return false;
        if (tempDoc.Blocks.Count != 1) return false;

        block = tempDoc.Blocks[0];
        index = index + local;
        return true;
    }

    private static List<IMarkdownBlock> ParseBlocksFromLines(string[] lines, MarkdownReaderOptions options, MarkdownReaderState state) {
        var doc = MarkdownDoc.Create();
        var opt = CloneOptionsWithoutFrontMatter(options);
        var pipeline = MarkdownReaderPipeline.Default(opt);

        int i = 0;
        while (i < lines.Length) {
            if (string.IsNullOrWhiteSpace(lines[i])) { i++; continue; }
            bool matched = false;
            var parsers = pipeline.Parsers;
            for (int p = 0; p < parsers.Count; p++) {
                if (parsers[p].TryParse(lines, ref i, opt, doc, state)) { matched = true; break; }
            }
            if (!matched) i++;
        }

        return doc.Blocks.ToList();
    }

    private static bool EndsWithTwoSpacesLine(string s) {
        if (string.IsNullOrEmpty(s)) return false;
        int n = s.Length - 1;
        int count = 0;
        while (n >= 0 && s[n] == ' ') {
            count++;
            n--;
            if (count >= 2) return true;
        }
        return false;
    }

    private static List<InlineSequence> ParseParagraphsFromLines(List<string> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        var paragraphs = new List<InlineSequence>();
        if (lines == null || lines.Count == 0) {
            paragraphs.Add(ParseInlines(string.Empty, options, state));
            return paragraphs;
        }

        var cur = new List<string>();
        for (int i = 0; i < lines.Count; i++) {
            var ln = lines[i] ?? string.Empty;
            if (ln.Length == 0) {
                if (cur.Count > 0) {
                    paragraphs.Add(ParseInlines(JoinParagraphLines(cur, options), options, state));
                    cur.Clear();
                }
                continue;
            }
            cur.Add(ln);
        }
        if (cur.Count > 0) paragraphs.Add(ParseInlines(JoinParagraphLines(cur, options), options, state));

        if (paragraphs.Count == 0) paragraphs.Add(ParseInlines(string.Empty, options, state));
        return paragraphs;
    }

    private static string JoinParagraphLines(List<string> lines, MarkdownReaderOptions options) {
        var sb = new StringBuilder();
        bool prevHard = false;
        for (int i = 0; i < lines.Count; i++) {
            var raw = lines[i] ?? string.Empty;
            bool hard = EndsWithTwoSpacesLine(raw);
            var trimmed = raw.TrimEnd();
            trimmed = ConsumeTrailingBackslashHardBreak(trimmed, options, out bool slashHard);
            hard = hard || slashHard;

            if (i > 0) sb.Append(prevHard ? "\n" : " ");
            sb.Append(trimmed);
            prevHard = hard;
        }
        return sb.ToString();
    }

    private static string ConsumeTrailingBackslashHardBreak(string trimmed, MarkdownReaderOptions options, out bool hardBreak) {
        hardBreak = false;
        if (options == null || !options.BackslashHardBreaks) return trimmed ?? string.Empty;
        if (string.IsNullOrEmpty(trimmed)) return string.Empty;
        if (trimmed[trimmed.Length - 1] != '\\') return trimmed;
        hardBreak = true;
        return trimmed.Substring(0, trimmed.Length - 1);
    }

    private static void ConsumeNestedBlocksForListItem(
        string[] lines,
        ref int index,
        int itemLevelAbs,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        ListItem item,
        bool allowNestedOrdered,
        bool allowNestedUnordered) {

        if (lines == null || item == null) return;

        int baseIndent = Math.Max(0, itemLevelAbs) * 2;
        int continuationIndent = baseIndent + 2;

        while (index < lines.Length) {
            int k = index;

            // Skip blank lines only when they are followed by nested content.
            while (k < lines.Length && string.IsNullOrWhiteSpace(lines[k])) {
                int peek = k + 1;
                if (peek >= lines.Length) return;
                var next = lines[peek] ?? string.Empty;
                int nextSpaces = CountLeadingSpaces(next);
                if (nextSpaces < continuationIndent) return;

                bool nestedStart = false;
                if (allowNestedOrdered && options.OrderedLists && IsOrderedListLine(next, out int lvlAbsO, out _, out _) && lvlAbsO >= itemLevelAbs + 1) nestedStart = true;
                if (allowNestedUnordered && options.UnorderedLists && IsUnorderedListLine(next, out int lvlAbsU, out _, out _, out _) && lvlAbsU >= itemLevelAbs + 1) nestedStart = true;
                if (options.FencedCode) {
                    var slice = next.Length >= continuationIndent ? next.Substring(continuationIndent) : next.TrimStart();
                    if (IsCodeFenceOpen(slice, out _, out _, out _)) nestedStart = true;
                }

                if (!nestedStart) return;
                k++;
            }

            if (k >= lines.Length) { index = k; return; }

            // Nested fenced code block
            int tmp = k;
            if (TryParseNestedFencedCodeBlock(lines, ref tmp, continuationIndent, options, out var code) && code != null) {
                item.Children.Add(code);
                index = tmp;
                continue;
            }

            // Nested indented code block
            tmp = k;
            if (TryParseNestedIndentedCodeBlock(lines, ref tmp, continuationIndent, options, out var indented) && indented != null) {
                item.Children.Add(indented);
                index = tmp;
                continue;
            }

            // Nested blockquote
            tmp = k;
            if (TryParseNestedQuoteBlock(lines, ref tmp, continuationIndent, options, state, out var quote) && quote != null) {
                item.Children.Add(quote);
                index = tmp;
                continue;
            }

            // Nested table
            tmp = k;
            if (TryParseNestedTableBlock(lines, ref tmp, continuationIndent, options, state, out var table) && table != null) {
                item.Children.Add(table);
                index = tmp;
                continue;
            }

            // Nested HTML blocks (details / raw HTML) when HtmlBlocks are enabled.
            tmp = k;
            if (TryParseNestedHtmlBlock(lines, ref tmp, continuationIndent, options, state, out var htmlBlock) && htmlBlock != null) {
                item.Children.Add(htmlBlock);
                index = tmp;
                continue;
            }

            // Nested ordered list
            if (allowNestedOrdered && options.OrderedLists && IsOrderedListLine(lines[k], out int lvlAbsO2, out _, out _) && lvlAbsO2 >= itemLevelAbs + 1) {
                int idx = k;
                var tempDoc = MarkdownDoc.Create();
                var parser = new OrderedListParser();
                if (parser.TryParse(lines, ref idx, options, tempDoc, state) && tempDoc.Blocks.Count == 1 && tempDoc.Blocks[0] is OrderedListBlock ol) {
                    item.Children.Add(ol);
                    index = idx;
                    continue;
                }
            }

            // Nested unordered list
            if (allowNestedUnordered && options.UnorderedLists && IsUnorderedListLine(lines[k], out int lvlAbsU2, out _, out _, out _) && lvlAbsU2 >= itemLevelAbs + 1) {
                int idx = k;
                var tempDoc = MarkdownDoc.Create();
                var parser = new UnorderedListParser();
                if (parser.TryParse(lines, ref idx, options, tempDoc, state) && tempDoc.Blocks.Count == 1 && tempDoc.Blocks[0] is UnorderedListBlock ul) {
                    item.Children.Add(ul);
                    index = idx;
                    continue;
                }
            }

            // Nothing nested to consume.
            index = k;
            return;
        }
    }

    private static bool IsDefinitionLine(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        var trimmed = line.TrimStart();
        if (IsAtxHeading(trimmed, out _, out _)) return false; // headings take priority over definition lists
        return TryGetDefinitionSeparator(line, out _);
    }

    private static bool TryGetDefinitionSeparator(string line, out int idx) {
        idx = -1;
        if (string.IsNullOrWhiteSpace(line)) return false;
        int start = 0;
        while (start < line.Length) {
            int pos = line.IndexOf(':', start);
            if (pos < 0) return false;
            if (pos > 0 && pos + 1 < line.Length && line[pos + 1] == ' ') {
                idx = pos;
                return true;
            }
            start = pos + 1;
        }
        return false;
    }

    private static bool IsOrderedListLine(string line, out int number, out string content) {
        number = 0; content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        // Allow indentation; compute after leading spaces
        int spaces = 0; while (spaces < line.Length && line[spaces] == ' ') spaces++;
        int i = spaces; while (i < line.Length && char.IsDigit(line[i])) i++;
        if (i == spaces) return false;
        if (i < line.Length && (line[i] == '.' || line[i] == ')') && i + 1 < line.Length && line[i + 1] == ' ') {
            if (!int.TryParse(line.Substring(spaces, i - spaces), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) number = 1;
            content = line.Substring(i + 2);
            return true;
        }
        return false;
    }

    private static bool IsOrderedListLine(string line, out int level, out int number, out string content) {
        level = 0; number = 0; content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        int spaces = 0; while (spaces < line.Length && line[spaces] == ' ') spaces++;
        int i = spaces; while (i < line.Length && char.IsDigit(line[i])) i++;
        if (i == spaces) return false;
        if (i < line.Length && (line[i] == '.' || line[i] == ')') && i + 1 < line.Length && line[i + 1] == ' ') {
            if (!int.TryParse(line.Substring(spaces, i - spaces), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) number = 1;
            content = line.Substring(i + 2);
            level = spaces / 2;
            return true;
        }
        return false;
    }

    private static bool IsUnorderedListLine(string line, out bool isTask, out bool done, out string content) {
        isTask = false; done = false; content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.TrimStart();
        if (t.StartsWith("- ") || t.StartsWith("* ") || t.StartsWith("+ ")) {
            var c = t.Substring(2);
            if (c.StartsWith("[ ]")) { isTask = true; done = false; content = c.Length > 3 && c[2] == ']' && c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c; return true; }
            if (c.StartsWith("[x]", StringComparison.OrdinalIgnoreCase)) { isTask = true; done = true; content = c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c; return true; }
            content = c; return true;
        }
        return false;
    }

    private static bool IsUnorderedListLine(string line, out int level, out bool isTask, out bool done, out string content) {
        level = 0; isTask = false; done = false; content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        int spaces = 0; while (spaces < line.Length && line[spaces] == ' ') spaces++;
        if (spaces >= line.Length) return false;
        char ch = line[spaces];
        if ((ch == '-' || ch == '*' || ch == '+') && spaces + 1 < line.Length && line[spaces + 1] == ' ') {
            string c = line.Substring(spaces + 2);
            if (c.StartsWith("[ ]")) { isTask = true; done = false; content = c.Length > 3 && c[2] == ']' && c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c; level = spaces / 2; return true; }
            if (c.StartsWith("[x]", StringComparison.OrdinalIgnoreCase)) { isTask = true; done = true; content = c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c; level = spaces / 2; return true; }
            content = c; level = spaces / 2; return true;
        }
        return false;
    }

    private static bool IsCalloutHeader(string line, out string kind, out string title) {
        kind = string.Empty; title = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.TrimStart();
        if (!t.StartsWith(">")) return false;
        t = t.Substring(1).TrimStart();
        if (!t.StartsWith("[!")) return false;
        int close = t.IndexOf(']');
        if (close < 0 || close < 3) return false;
        string marker = t.Substring(2, close - 2);
        for (int i = 0; i < marker.Length; i++) if (!char.IsLetter(marker[i])) return false;
        kind = marker.ToLowerInvariant();
        title = t.Substring(close + 1).TrimStart();
        // Title is optional: "> [!NOTE]" is valid and should produce a callout with the default title for the kind.
        return true;
    }

    private static Dictionary<string, object?> ParseFrontMatter(string[] lines, int start, int end) {
        var dict = new Dictionary<string, object?>(StringComparer.OrdinalIgnoreCase);
        for (int i = start; i <= end; i++) {
            var line = lines[i]; if (string.IsNullOrWhiteSpace(line)) continue;
            int idx = line.IndexOf(':'); if (idx <= 0) continue;
            string key = line.Substring(0, idx).Trim(); string val = line.Substring(idx + 1).TrimStart();
            if (val == "|") {
                var sb = new StringBuilder(); int j = i + 1;
                while (j <= end) { var raw = lines[j]; if (raw.StartsWith("  ")) { sb.AppendLine(raw.Substring(2)); j++; } else break; }
                i = j - 1; dict[key] = sb.ToString().TrimEnd(); continue;
            }
            if (val.StartsWith("[") && val.EndsWith("]")) {
                var inner = val.Substring(1, val.Length - 2).Trim(); var items = new List<string>(); var token = new StringBuilder(); bool inQuotes = false;
                for (int k = 0; k < inner.Length; k++) { char ch = inner[k]; if (ch == '\"') { inQuotes = !inQuotes; continue; } if (ch == ',' && !inQuotes) { items.Add(token.ToString().Trim()); token.Clear(); continue; } token.Append(ch); }
                if (token.Length > 0) items.Add(token.ToString().Trim());
                dict[key] = items;
            } else if (string.Equals(val, "true", StringComparison.OrdinalIgnoreCase)) { dict[key] = true; } else if (string.Equals(val, "false", StringComparison.OrdinalIgnoreCase)) { dict[key] = false; } else if (double.TryParse(val, NumberStyles.Any, CultureInfo.InvariantCulture, out var num)) { dict[key] = num; } else if (val.StartsWith("\"") && val.EndsWith("\"")) { dict[key] = val.Length >= 2 ? val.Substring(1, val.Length - 2) : string.Empty; } else { dict[key] = val; }
        }
        return dict;
    }
}
