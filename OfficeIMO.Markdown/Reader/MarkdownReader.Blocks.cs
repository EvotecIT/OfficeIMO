using System.Globalization;
using System.Linq;

namespace OfficeIMO.Markdown;

/// <summary>
/// Block parsing helpers for <see cref="MarkdownReader"/>.
/// </summary>
public static partial class MarkdownReader {
    private static bool IsAtxHeading(string line, out int level, out string text) {
        return TryGetAtxHeadingContentRange(line, out level, out _, out _, out text);
    }

    private static bool TryGetAtxHeadingContentRange(string line, out int level, out int contentStart, out int contentEnd, out string text) {
        level = 0;
        contentStart = 0;
        contentEnd = 0;
        text = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;

        int indent = 0;
        while (indent < line.Length && indent < 4 && line[indent] == ' ') indent++;
        if (indent >= 4) return false;

        int i = indent;
        while (i < line.Length && line[i] == '#') i++;

        int count = i - indent;
        if (count < 1 || count > 6) return false;
        if (i < line.Length && !char.IsWhiteSpace(line[i])) return false;

        contentStart = i;
        while (contentStart < line.Length && char.IsWhiteSpace(line[contentStart])) contentStart++;
        if (contentStart >= line.Length) {
            level = count;
            text = string.Empty;
            contentEnd = contentStart;
            return true;
        }

        contentEnd = line.Length;
        while (contentEnd > contentStart && char.IsWhiteSpace(line[contentEnd - 1])) contentEnd--;

        int closingStart = contentEnd;
        while (closingStart > contentStart && line[closingStart - 1] == '#') closingStart--;
        if (closingStart < contentEnd) {
            int beforeClosing = closingStart - 1;
            if (beforeClosing >= contentStart && char.IsWhiteSpace(line[beforeClosing])) {
                contentEnd = beforeClosing;
                while (contentEnd > contentStart && char.IsWhiteSpace(line[contentEnd - 1])) contentEnd--;
            }
        }

        level = count;
        text = line.Substring(contentStart, contentEnd - contentStart);
        return true;
    }

    private static bool IsCodeFenceOpen(string line, out string infoString, out char fenceChar, out int fenceLength) {
        infoString = string.Empty; fenceChar = '\0'; fenceLength = 0;
        if (line is null) return false;
        int indent = CountLeadingIndentColumns(line);
        if (indent > 3) return false;

        line = indent > 0 ? StripLeadingIndentColumns(line, indent) : line;
        if (line.Length < 3) return false;
        char ch = line[0];
        if (ch != '`' && ch != '~') return false;

        int run = 0;
        while (run < line.Length && line[run] == ch) run++;
        if (run < 3) return false;

        var parsedInfoString = line.Length > run ? line.Substring(run) : string.Empty;
        if (ch == '`' && parsedInfoString.IndexOf('`') >= 0) return false;

        fenceChar = ch;
        fenceLength = run;
        infoString = parsedInfoString.Trim();
        return true;
    }
    private static bool IsCodeFenceClose(string line, char fenceChar, int fenceLength) {
        if (line is null) return false;
        int indent = CountLeadingIndentColumns(line);
        if (indent > 3) return false;

        var trimmed = (indent > 0 ? StripLeadingIndentColumns(line, indent) : line).Trim();
        if (trimmed.Length < Math.Max(3, fenceLength)) return false;
        // CommonMark allows closing fence length >= opening fence length. We accept that.
        for (int i = 0; i < trimmed.Length; i++) {
            if (trimmed[i] != fenceChar) return false;
        }
        return trimmed.Length >= Math.Max(3, fenceLength);
    }

    private static IMarkdownBlock CreateParsedFencedBlock(string infoString, string content, bool isFenced, string? caption, MarkdownReaderOptions options) {
        var extendedBlock = TryCreateExtendedFencedBlock(options?.FencedBlockExtensions, infoString, content, isFenced, caption);
        if (extendedBlock != null) {
            return extendedBlock;
        }

        return new CodeBlock(infoString, content, isFenced) {
            Caption = caption
        };
    }

    internal static IMarkdownBlock? TryCreateExtendedFencedBlock(
        IReadOnlyList<MarkdownFencedBlockExtension>? extensions,
        string infoString,
        string content,
        bool isFenced,
        string? caption) {
        if (extensions == null || extensions.Count == 0) {
            return null;
        }

        var context = new MarkdownFencedBlockFactoryContext(infoString, content, isFenced, caption);
        for (int i = extensions.Count - 1; i >= 0; i--) {
            var extension = extensions[i];
            if (extension == null || !FencedBlockExtensionHandlesLanguage(extension, context.Language)) {
                continue;
            }

            var block = extension.CreateBlock(context);
            if (block == null) {
                continue;
            }

            if (!string.IsNullOrWhiteSpace(caption) && block is ICaptionable captionable && string.IsNullOrWhiteSpace(captionable.Caption)) {
                captionable.Caption = caption;
            }

            return block;
        }

        return null;
    }

    private static string RemoveSingleTrailingLineEnding(string text) {
        if (string.IsNullOrEmpty(text)) {
            return string.Empty;
        }

        if (text.EndsWith("\r\n", StringComparison.Ordinal)) {
            return text.Substring(0, text.Length - 2);
        }

        if (text[text.Length - 1] == '\n' || text[text.Length - 1] == '\r') {
            return text.Substring(0, text.Length - 1);
        }

        return text;
    }

    private static bool FencedBlockExtensionHandlesLanguage(MarkdownFencedBlockExtension extension, string language) {
        var languages = extension.Languages;
        if (languages == null || languages.Count == 0) {
            return false;
        }

        for (int i = 0; i < languages.Count; i++) {
            var candidate = languages[i];
            if (!string.IsNullOrWhiteSpace(candidate) && string.Equals(candidate, language, StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private static bool TryParseCaption(string line, out string caption) {
        caption = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.Trim();
        if (t.Length >= 3 && t[0] == '_' && t[t.Length - 1] == '_' && t.IndexOf('_', 1) == t.Length - 1) { caption = t.Substring(1, t.Length - 2); return true; }
        return false;
    }

    private readonly struct MarkdownImageSyntaxRanges(
        int? altStart,
        int? altLength,
        int? sourceStart,
        int? sourceLength,
        int? titleStart,
        int? titleLength,
        int? linkTargetStart,
        int? linkTargetLength,
        int? linkTitleStart,
        int? linkTitleLength) {
        public int? AltStart { get; } = altStart;
        public int? AltLength { get; } = altLength;
        public int? SourceStart { get; } = sourceStart;
        public int? SourceLength { get; } = sourceLength;
        public int? TitleStart { get; } = titleStart;
        public int? TitleLength { get; } = titleLength;
        public int? LinkTargetStart { get; } = linkTargetStart;
        public int? LinkTargetLength { get; } = linkTargetLength;
        public int? LinkTitleStart { get; } = linkTitleStart;
        public int? LinkTitleLength { get; } = linkTitleLength;
    }

    private static bool IsImageLine(string line) {
        ImageBlock image;
        string? sizeSpec;
        MarkdownImageSyntaxRanges ranges;
        return TryParseImage(line, out image, out sizeSpec, out ranges);
    }
    private static bool TryParseImage(string line, out ImageBlock image, out string? sizeSpec) =>
        TryParseImage(line, new MarkdownReaderOptions(), new MarkdownReaderState(), out image, out sizeSpec, out _);

    private static bool TryParseImage(string line, out ImageBlock image, out string? sizeSpec, out MarkdownImageSyntaxRanges ranges) =>
        TryParseImage(line, new MarkdownReaderOptions(), new MarkdownReaderState(), out image, out sizeSpec, out ranges);

    private static bool TryParseImage(string line, MarkdownReaderOptions options, MarkdownReaderState state, out ImageBlock image, out string? sizeSpec) =>
        TryParseImage(line, options, state, out image, out sizeSpec, out _);

    private static bool TryParseImage(string line, MarkdownReaderOptions options, MarkdownReaderState state, out ImageBlock image, out string? sizeSpec, out MarkdownImageSyntaxRanges ranges) {
        image = null!;
        sizeSpec = null;
        ranges = default;
        if (string.IsNullOrEmpty(line)) return false;
        var t = line.Trim();
        if (!t.StartsWith("![", StringComparison.Ordinal)) return false;
        int altEnd = FindMatchingBracket(t, 1);
        if (altEnd < 2) return false;
        if (altEnd + 1 >= t.Length || t[altEnd + 1] != '(') return false;
        int parenClose = FindMatchingParen(t, altEnd + 1);
        if (parenClose <= altEnd + 2) return false;
        string alt = t.Substring(2, altEnd - 2);
        string inner = t.Substring(altEnd + 2, parenClose - (altEnd + 2));
        if (!TrySplitUrlAndOptionalTitle(inner, out var src, out var title, out int srcStart, out int srcLength, out int? titleStart, out int? titleLength)) {
            if (IndexOfWhitespace(inner.Trim()) >= 0) return false;
            src = UnescapeMarkdownBackslashEscapes(inner.Trim());
            title = null;
            int trimmedStart = 0;
            while (trimmedStart < inner.Length && char.IsWhiteSpace(inner[trimmedStart])) {
                trimmedStart++;
            }

            int trimmedEndExclusive = inner.Length;
            while (trimmedEndExclusive > trimmedStart && char.IsWhiteSpace(inner[trimmedEndExclusive - 1])) {
                trimmedEndExclusive--;
            }

            srcStart = trimmedStart;
            srcLength = Math.Max(0, trimmedEndExclusive - trimmedStart);
            titleStart = null;
            titleLength = null;
        }
        string plainAlt = ExtractImageAltPlainText(alt, options, state);
        image = new ImageBlock(src, alt, title, plainAlt: plainAlt);
        ranges = new MarkdownImageSyntaxRanges(
            altStart: 2,
            altLength: altEnd - 2,
            sourceStart: altEnd + 2 + srcStart,
            sourceLength: srcLength,
            titleStart: titleStart.HasValue ? altEnd + 2 + titleStart.Value : null,
            titleLength: titleLength,
            linkTargetStart: null,
            linkTargetLength: null,
            linkTitleStart: null,
            linkTitleLength: null);
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

    private static bool TryParseLinkedImageBlock(string line, out ImageBlock image, out string? sizeSpec) =>
        TryParseLinkedImageBlock(line, new MarkdownReaderOptions(), new MarkdownReaderState(), out image, out sizeSpec, out _);

    private static bool TryParseLinkedImageBlock(string line, out ImageBlock image, out string? sizeSpec, out MarkdownImageSyntaxRanges ranges) =>
        TryParseLinkedImageBlock(line, new MarkdownReaderOptions(), new MarkdownReaderState(), out image, out sizeSpec, out ranges);

    private static bool TryParseLinkedImageBlock(string line, MarkdownReaderOptions options, MarkdownReaderState state, out ImageBlock image, out string? sizeSpec) =>
        TryParseLinkedImageBlock(line, options, state, out image, out sizeSpec, out _);

    private static bool TryParseLinkedImageBlock(string line, MarkdownReaderOptions options, MarkdownReaderState state, out ImageBlock image, out string? sizeSpec, out MarkdownImageSyntaxRanges ranges) {
        image = null!;
        sizeSpec = null;
        ranges = default;
        if (string.IsNullOrEmpty(line)) {
            return false;
        }

        var t = line.Trim();
        if (!TryParseImageLink(
            t,
            0,
            out int consumed,
            out var alt,
            out var src,
            out var title,
            out var href,
            out var hrefTitle,
            out int altStart,
            out int altLength,
            out int srcStart,
            out int srcLength,
            out int? titleStart,
            out int? titleLength,
            out int hrefStart,
            out int hrefLength,
            out int? hrefTitleStart,
            out int? hrefTitleLength) || consumed <= 0) {
            return false;
        }

        string plainAlt = ExtractImageAltPlainText(alt, options, state);
        image = new ImageBlock(src, alt, title, plainAlt: plainAlt) {
            LinkUrl = href,
            LinkTitle = hrefTitle
        };
        ranges = new MarkdownImageSyntaxRanges(
            altStart,
            altLength,
            srcStart,
            srcLength,
            titleStart,
            titleLength,
            hrefStart,
            hrefLength,
            hrefTitleStart,
            hrefTitleLength);

        if (consumed < t.Length) {
            var rest = t.Substring(consumed).Trim();
            if (rest.StartsWith("{", StringComparison.Ordinal)) {
                int close = rest.IndexOf('}');
                if (close > 0) {
                    sizeSpec = rest.Substring(1, close - 1).Trim();
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

    private static TableBlock ParseTable(string[] lines, int start, int end, MarkdownReaderOptions options, MarkdownReaderState state) {
        int headerLine = state.SourceLineOffset + start + 1;
        var cells0 = SplitTableRowWithSourceInfo(lines[start], headerLine, state);
        var table = new TableBlock();
        var inlineOptions = CloneOptionsWithoutFrontMatter(options);
        var inlineState = CloneState(state);
        table.InlineRenderOptions = inlineOptions;
        table.InlineRenderState = inlineState;
        if (start + 1 <= end && IsAlignmentRow(lines[start + 1])) {
            table.Headers.AddRange(cells0.Select(cell => cell.Text));
            var aligns = SplitTableRow(lines[start + 1]);
            for (int i = 0; i < aligns.Count; i++) table.Alignments.Add(ParseAlignmentCell(aligns[i]));
            for (int i = start + 2; i <= end; i++) {
                int absoluteLine = state.SourceLineOffset + i + 1;
                var row = SplitTableRowWithSourceInfo(lines[i], absoluteLine, state);
                table.Rows.Add(row.Select(cell => cell.Text).ToArray());
            }
        } else {
            for (int i = start; i <= end; i++) {
                int absoluteLine = state.SourceLineOffset + i + 1;
                var row = SplitTableRowWithSourceInfo(lines[i], absoluteLine, state);
                table.Rows.Add(row.Select(cell => cell.Text).ToArray());
            }
        }

        var headerCells = table.Headers.Count > 0
            ? SplitTableRowWithSourceInfo(lines[start], headerLine, state)
            : new List<TableCellSourceFragment>();
        var bodyRows = new List<IReadOnlyList<TableCellSourceFragment>>(table.Rows.Count);
        if (table.Headers.Count > 0) {
            for (int i = start + 2; i <= end; i++) {
                int absoluteLine = state.SourceLineOffset + i + 1;
                bodyRows.Add(SplitTableRowWithSourceInfo(lines[i], absoluteLine, state));
            }
        } else {
            for (int i = start; i <= end; i++) {
                int absoluteLine = state.SourceLineOffset + i + 1;
                bodyRows.Add(SplitTableRowWithSourceInfo(lines[i], absoluteLine, state));
            }
        }

        table.SetParsedCells(
            ParseTableInlineCells(headerCells, inlineOptions, inlineState),
            ParseTableInlineRows(bodyRows, inlineOptions, inlineState),
            table.ComputeContentSignature());
        table.SetStructuredCells(
            BuildTableCells(headerCells, inlineOptions, inlineState),
            BuildTableRows(bodyRows, inlineOptions, inlineState),
            table.ComputeContentSignature());
        return table;
    }

    private static List<InlineSequence> ParseTableInlineCells(IReadOnlyList<TableCellSourceFragment> cells, MarkdownReaderOptions options, MarkdownReaderState state) {
        var parsedCells = new List<InlineSequence>(cells?.Count ?? 0);
        if (cells == null) {
            return parsedCells;
        }

        for (int i = 0; i < cells.Count; i++) {
            parsedCells.Add(ParseTableCellInlines(cells[i], options, state));
        }
        return parsedCells;
    }

    private static List<IReadOnlyList<InlineSequence>> ParseTableInlineRows(IReadOnlyList<IReadOnlyList<TableCellSourceFragment>> rows, MarkdownReaderOptions options, MarkdownReaderState state) {
        var parsedRows = new List<IReadOnlyList<InlineSequence>>(rows.Count);
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            if (row == null || row.Count == 0) {
                parsedRows.Add(Array.Empty<InlineSequence>());
                continue;
            }

            var parsedRow = new List<InlineSequence>(row.Count);
            for (int cellIndex = 0; cellIndex < row.Count; cellIndex++) {
                parsedRow.Add(ParseTableCellInlines(row[cellIndex], options, state));
            }
            parsedRows.Add(parsedRow);
        }
        return parsedRows;
    }

    private static InlineSequence ParseTableCellInlines(TableCellSourceFragment cell, MarkdownReaderOptions options, MarkdownReaderState state) {
        if (string.IsNullOrEmpty(cell.Markdown)) {
            return new InlineSequence();
        }

        var (sanitized, sourceMap) = NormalizeTableCellInlineMarkdownWithSourceMap(cell, state);
        return ParseInlines(sanitized, options, state, sourceMap);
    }

    private static List<TableCell> BuildTableCells(IReadOnlyList<TableCellSourceFragment> cells, MarkdownReaderOptions options, MarkdownReaderState state) {
        var typedCells = new List<TableCell>(cells?.Count ?? 0);
        if (cells == null) {
            return typedCells;
        }

        for (int i = 0; i < cells.Count; i++) {
            typedCells.Add(BuildTableCell(cells[i], options, state));
        }
        return typedCells;
    }

    private static List<IReadOnlyList<TableCell>> BuildTableRows(IReadOnlyList<IReadOnlyList<TableCellSourceFragment>> rows, MarkdownReaderOptions options, MarkdownReaderState state) {
        var typedRows = new List<IReadOnlyList<TableCell>>(rows.Count);
        for (int rowIndex = 0; rowIndex < rows.Count; rowIndex++) {
            var row = rows[rowIndex];
            if (row == null || row.Count == 0) {
                typedRows.Add(Array.Empty<TableCell>());
                continue;
            }

            typedRows.Add(BuildTableCells(row, options, state));
        }
        return typedRows;
    }

    private static TableCell BuildTableCell(TableCellSourceFragment cell, MarkdownReaderOptions options, MarkdownReaderState state) {
        if (string.IsNullOrEmpty(cell.Markdown)) {
            return new TableCell();
        }

        var structuredCell = TryParseStructuredTableCellBlocks(cell, options, state);
        if (structuredCell.HasValue) {
            return new TableCell(structuredCell.Value.Blocks) {
                SourceSpan = cell.SourceSpan,
                SyntaxChildren = structuredCell.Value.SyntaxChildren
            };
        }

        return new TableCell(new[] {
            new ParagraphBlock(ParseTableCellInlines(cell, options, state))
        }) {
            SourceSpan = cell.SourceSpan
        };
    }

    private static (IReadOnlyList<IMarkdownBlock> Blocks, IReadOnlyList<MarkdownSyntaxNode> SyntaxChildren)? TryParseStructuredTableCellBlocks(
        TableCellSourceFragment cell,
        MarkdownReaderOptions options,
        MarkdownReaderState state) {
        if (string.IsNullOrEmpty(cell.Markdown)) {
            return null;
        }

        var normalized = BuildTableCellSourceLines(cell);
        if (!TableBlock.LooksLikeStructuredMarkdownCell(string.Join("\n", normalized.Select(line => line.Text)))) {
            return null;
        }

        var fragmentOptions = CloneOptionsWithoutFrontMatter(options);
        fragmentOptions.Tables = false;
        var fragmentState = CloneState(state);
        var (blocks, syntaxChildren) = ParseNestedMarkdownBlocks(normalized, fragmentOptions, fragmentState);
        if (blocks.Count == 0) {
            return null;
        }

        if (TableBlock.ContainsUnsafeRawHtmlTableCellBlocks(blocks)) {
            return null;
        }

        if (blocks.Count == 1 && blocks[0] is ParagraphBlock) {
            return null;
        }

        return (blocks, syntaxChildren);
    }

    private readonly struct TableCellSourceFragment {
        public TableCellSourceFragment(string markdown, string text, MarkdownSourceSpan sourceSpan) {
            Markdown = markdown ?? string.Empty;
            Text = text ?? string.Empty;
            SourceSpan = sourceSpan;
        }

        public string Markdown { get; }
        public string Text { get; }
        public MarkdownSourceSpan SourceSpan { get; }
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
        return SplitTableRowWithSourceInfo(line, absoluteLine: 1, state: null).Select(cell => cell.Text).ToList();
    }

    private static List<TableCellSourceFragment> SplitTableRowWithSourceInfo(string line, int absoluteLine, MarkdownReaderState? state) {
        if (line is null) {
            return new List<TableCellSourceFragment>();
        }

        int trimStart = 0;
        int trimEndExclusive = line.Length;
        while (trimStart < trimEndExclusive && char.IsWhiteSpace(line[trimStart])) {
            trimStart++;
        }
        while (trimEndExclusive > trimStart && char.IsWhiteSpace(line[trimEndExclusive - 1])) {
            trimEndExclusive--;
        }

        int contentStart = trimStart;
        int contentEndExclusive = trimEndExclusive;
        if (contentStart < contentEndExclusive && line[contentStart] == '|') {
            contentStart++;
        }
        if (contentEndExclusive > contentStart && line[contentEndExclusive - 1] == '|') {
            contentEndExclusive--;
        }

        var cells = new List<TableCellSourceFragment>();
        int segmentStart = contentStart;
        int index = contentStart;
        int codeFenceLen = 0;

        while (index < contentEndExclusive) {
            char ch = line[index];
            if (ch == '\\' && index + 1 < contentEndExclusive) {
                index += 2;
                continue;
            }

            if (ch == '`') {
                int run = 1;
                int tick = index + 1;
                while (tick < contentEndExclusive && line[tick] == '`') {
                    run++;
                    tick++;
                }

                if (codeFenceLen == 0) {
                    if (HasClosingBacktickRun(line.Substring(contentStart, contentEndExclusive - contentStart), index - contentStart + run, run)) {
                        codeFenceLen = run;
                    }
                } else if (run == codeFenceLen) {
                    codeFenceLen = 0;
                }

                index += run;
                continue;
            }

            if (ch == '|' && codeFenceLen == 0) {
                cells.Add(CreateTableCellSourceFragment(line, segmentStart, index, absoluteLine, state));
                segmentStart = index + 1;
            }

            index++;
        }

        cells.Add(CreateTableCellSourceFragment(line, segmentStart, contentEndExclusive, absoluteLine, state));
        return cells;
    }

    private static TableCellSourceFragment CreateTableCellSourceFragment(
        string line,
        int segmentStart,
        int segmentEndExclusive,
        int absoluteLine,
        MarkdownReaderState? state) {
        int trimmedStart = segmentStart;
        int trimmedEndExclusive = segmentEndExclusive;

        while (trimmedStart < trimmedEndExclusive && char.IsWhiteSpace(line[trimmedStart])) {
            trimmedStart++;
        }

        while (trimmedEndExclusive > trimmedStart && char.IsWhiteSpace(line[trimmedEndExclusive - 1])) {
            trimmedEndExclusive--;
        }

        string markdown = trimmedStart < trimmedEndExclusive
            ? line.Substring(trimmedStart, trimmedEndExclusive - trimmedStart)
            : string.Empty;
        string text = UnescapeBackslashEscapesOutsideCodeSpans(markdown);
        int startColumn = trimmedStart + 1;
        int endColumn = trimmedStart < trimmedEndExclusive ? trimmedEndExclusive : startColumn;
        var span = CreateSpan(state, absoluteLine, startColumn, absoluteLine, endColumn);
        return new TableCellSourceFragment(markdown, text, span);
    }

    private static (string Text, MarkdownInlineSourceMap? SourceMap) NormalizeTableCellInlineMarkdownWithSourceMap(
        TableCellSourceFragment cell,
        MarkdownReaderState? state) {
        if (string.IsNullOrEmpty(cell.Markdown)) {
            return (string.Empty, null);
        }

        var builder = new StringBuilder(cell.Markdown.Length + 8);
        List<MarkdownSourcePoint?>? points = state?.SourceTextMap == null ? null : new List<MarkdownSourcePoint?>(cell.Markdown.Length + 8);
        int absoluteLine = cell.SourceSpan.StartLine;
        int column = cell.SourceSpan.StartColumn ?? 1;

        void AppendMapped(string value, int sourceColumn) {
            builder.Append(value);
            if (points == null) {
                return;
            }

            var point = state!.SourceTextMap!.CreatePoint(absoluteLine, sourceColumn);
            for (int i = 0; i < value.Length; i++) {
                points.Add(point);
            }
        }

        for (int i = 0; i < cell.Markdown.Length; i++) {
            char ch = cell.Markdown[i];
            if (ch == '\r') {
                if (i + 1 < cell.Markdown.Length && cell.Markdown[i + 1] == '\n') {
                    i++;
                }
                AppendMapped("\n", column);
                column++;
                continue;
            }

            if (ch == '\n') {
                AppendMapped("\n", column);
                column++;
                continue;
            }

            if (ch == '<' && TableBlock.TryConsumeBreakTag(cell.Markdown, i, out int consumed)) {
                AppendMapped("\n", column);
                column += consumed;
                i += consumed - 1;
                continue;
            }

            switch (ch) {
                case '<':
                    AppendMapped("&lt;", column);
                    break;
                case '>':
                    AppendMapped("&gt;", column);
                    break;
                case '&':
                    AppendMapped("&amp;", column);
                    break;
                default:
                    AppendMapped(ch.ToString(), column);
                    break;
            }

            column++;
        }

        return (builder.ToString(), points == null ? null : new MarkdownInlineSourceMap(points.ToArray()));
    }

    private static List<MarkdownSourceLineSlice> BuildTableCellSourceLines(TableCellSourceFragment cell) {
        var slices = new List<MarkdownSourceLineSlice>();
        if (string.IsNullOrEmpty(cell.Markdown)) {
            slices.Add(new MarkdownSourceLineSlice(string.Empty, cell.SourceSpan.StartLine, cell.SourceSpan.StartColumn ?? 1));
            return slices;
        }

        var current = new StringBuilder(cell.Markdown.Length);
        int absoluteLine = cell.SourceSpan.StartLine;
        int column = cell.SourceSpan.StartColumn ?? 1;
        int currentStartColumn = column;

        void Flush() {
            slices.Add(new MarkdownSourceLineSlice(current.ToString(), absoluteLine, currentStartColumn));
            current.Clear();
            currentStartColumn = column;
        }

        for (int i = 0; i < cell.Markdown.Length; i++) {
            char ch = cell.Markdown[i];
            if (ch == '\r') {
                if (i + 1 < cell.Markdown.Length && cell.Markdown[i + 1] == '\n') {
                    i++;
                }
                Flush();
                column++;
                currentStartColumn = column;
                continue;
            }

            if (ch == '\n') {
                Flush();
                column++;
                currentStartColumn = column;
                continue;
            }

            if (ch == '<' && TableBlock.TryConsumeBreakTag(cell.Markdown, i, out int consumed)) {
                Flush();
                column += consumed;
                currentStartColumn = column;
                i += consumed - 1;
                continue;
            }

            current.Append(ch);
            column++;
        }

        Flush();
        return slices;
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

                if (codeFenceLen == 0) {
                    if (HasClosingBacktickRun(value, j, run)) {
                        codeFenceLen = run;
                    }
                }
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

    private static bool HasClosingBacktickRun(string text, int start, int runLength) {
        if (string.IsNullOrEmpty(text) || start >= text.Length) return false;

        for (int i = start; i < text.Length; i++) {
            if (text[i] != '`') continue;

            int run = 1;
            int j = i + 1;
            while (j < text.Length && text[j] == '`') {
                run++;
                j++;
            }

            if (run == runLength) return true;
            i = j - 1;
        }

        return false;
    }

    private static int CountLeadingSpaces(string line) {
        if (string.IsNullOrEmpty(line)) return 0;
        int i = 0;
        while (i < line.Length && line[i] == ' ') i++;
        return i;
    }

    private static int CountLeadingIndentColumns(string line) {
        if (string.IsNullOrEmpty(line)) return 0;

        int columns = 0;
        for (int i = 0; i < line.Length; i++) {
            char ch = line[i];
            if (ch == ' ') {
                columns++;
                continue;
            }

            if (ch == '\t') {
                columns += 4 - (columns % 4);
                continue;
            }

            break;
        }

        return columns;
    }

    private static string StripLeadingIndentColumns(string line, int requiredColumns) {
        if (string.IsNullOrEmpty(line) || requiredColumns <= 0) return line ?? string.Empty;

        int columns = 0;
        int index = 0;
        while (index < line.Length && columns < requiredColumns) {
            char ch = line[index];
            if (ch == ' ') {
                columns++;
                index++;
                continue;
            }

            if (ch == '\t') {
                columns += 4 - (columns % 4);
                index++;
                continue;
            }

            break;
        }

        return line.Substring(index);
    }

    private static bool IsParagraphInterruptingOrderedListLine(string line) {
        if (!IsOrderedListLine(line, out _, out int number, out string content)) return false;
        return number == 1 && !string.IsNullOrWhiteSpace(content);
    }

    private static bool IsParagraphInterruptingUnorderedListLine(string line) {
        return IsUnorderedListLine(line, out _, out _, out string content)
               && !string.IsNullOrWhiteSpace(content);
    }

    private static bool LastCollectedLinePreservesIndentedContinuation(List<string> collected) {
        if (collected == null || collected.Count == 0) return false;

        for (int i = collected.Count - 1; i >= 0; i--) {
            var line = collected[i];
            if (string.IsNullOrWhiteSpace(line)) continue;
            if (!IsOrderedListLine(line, out _, out int number, out _)) return false;
            return number != 1;
        }

        return false;
    }

    private static List<string> ConsumeListContinuationLines(
        string[] lines,
        ref int nextIndex,
        int continuationIndent,
        string initialContent,
        MarkdownReaderOptions options,
        bool breakOnAnyOrderedListLine = false,
        List<MarkdownSourceLineSlice>? sourceLines = null,
        int absoluteLineOffset = 0,
        int initialLineIndex = -1,
        int initialStartColumn = 1) {
        if (lines == null) return new List<string> { initialContent ?? string.Empty };
        if (nextIndex < 0) nextIndex = 0;

        var collected = new List<string> { initialContent ?? string.Empty };
        if (sourceLines != null) {
            int initialAbsoluteLine = initialLineIndex >= 0
                ? absoluteLineOffset + initialLineIndex + 1
                : absoluteLineOffset + nextIndex + 1;
            sourceLines.Add(new MarkdownSourceLineSlice(initialContent ?? string.Empty, initialAbsoluteLine, initialStartColumn));
        }

        int k = nextIndex;

        while (k < lines.Length) {
            var line = lines[k] ?? string.Empty;
            bool collectingLeadFencedCode = TryGetOpenLeadFencedCode(collected, out _, out _, out _);

            if (collectingLeadFencedCode) {
                if (string.IsNullOrWhiteSpace(line)) {
                    collected.Add(string.Empty);
                    sourceLines?.Add(new MarkdownSourceLineSlice(string.Empty, absoluteLineOffset + k + 1, 1));
                    k++;
                    continue;
                }

                int fencedIndentColumns = CountLeadingIndentColumns(line);
                if (fencedIndentColumns < continuationIndent) {
                    break;
                }

                string fencedContent = StripLeadingIndentColumns(line, continuationIndent);
                int fencedStartColumn = continuationIndent + 1;
                sourceLines?.Add(new MarkdownSourceLineSlice(fencedContent, absoluteLineOffset + k + 1, fencedStartColumn));
                collected.Add(fencedContent);
                k++;
                continue;
            }

            // Stop before the next list item (including nested items).
            if (IsUnorderedListLine(line, out _, out _, out _, out _) ||
                (breakOnAnyOrderedListLine ? IsOrderedListLine(line, out _, out _, out _) : IsParagraphInterruptingOrderedListLine(line))) {
                break;
            }

            // Stop before nested blocks; they are handled as child blocks of the list item.
            if (CountLeadingIndentColumns(line) >= continuationIndent) {
                var slice = StripLeadingIndentColumns(line, continuationIndent);
                var sliceTrim = slice.TrimStart();
                if (IsCodeFenceOpen(slice, out _, out _, out _)) break;
                if (sliceTrim.StartsWith(">")) break;

                if (options.HtmlBlocks && sliceTrim.StartsWith("<")) {
                    // Avoid breaking on angle-bracket autolinks like "<https://...>".
                    if (!TryParseAngleAutolink(sliceTrim, 0, out _, out _, out _)) break;
                }

                // Indented code block inside list item: continuationIndent + 4 spaces.
                if (options.IndentedCodeBlocks) {
                    int lineIndentColumns = CountLeadingIndentColumns(line);
                    if (lineIndentColumns >= continuationIndent + 4 && !LastCollectedLinePreservesIndentedContinuation(collected)) break;
                }

                // Table inside list item: a pipe row followed by an alignment/row.
                if (options.Tables && LooksLikeTableRow(sliceTrim)) {
                    int peek = k + 1;
                    if (peek < lines.Length && CountLeadingIndentColumns(lines[peek] ?? string.Empty) >= continuationIndent) {
                        var nextSlice = StripLeadingIndentColumns(lines[peek] ?? string.Empty, continuationIndent).TrimStart();
                        // Reduce false positives: require an alignment row, or explicit outer pipes on both rows.
                        bool curOuter = sliceTrim.Length > 0 && sliceTrim[0] == '|' && sliceTrim[sliceTrim.Length - 1] == '|';
                        bool nextOuter = nextSlice.Length > 0 && nextSlice[0] == '|' && nextSlice[nextSlice.Length - 1] == '|';
                        if (IsAlignmentRow(nextSlice) || (curOuter && nextOuter)) break;
                    }
                }
            }

            if (string.IsNullOrWhiteSpace(line)) {
                if (collected.Count == 1 && string.IsNullOrWhiteSpace(collected[0])) {
                    break;
                }

                // Keep blank lines only if followed by an indented continuation line; otherwise end item.
                int peek = k + 1;
                if (peek >= lines.Length) break;
                var next = lines[peek] ?? string.Empty;
                if (IsUnorderedListLine(next, out _, out _, out _, out _) ||
                    (breakOnAnyOrderedListLine ? IsOrderedListLine(next, out _, out _, out _) : IsParagraphInterruptingOrderedListLine(next))) {
                    break;
                }
                int nextIndentColumns = CountLeadingIndentColumns(next);
                if (nextIndentColumns < continuationIndent) break;

                collected.Add(string.Empty);
                sourceLines?.Add(new MarkdownSourceLineSlice(string.Empty, absoluteLineOffset + k + 1, 1));
                k++;
                continue;
            }

            int indentColumns = CountLeadingIndentColumns(line);
            if (indentColumns < continuationIndent) {
                if (collected.Count > 0 &&
                    !string.IsNullOrWhiteSpace(collected[collected.Count - 1]) &&
                    LooksLikeParagraphLine(collected, collected.Count - 1, options) &&
                    TryNormalizeListLazyContinuationLine(lines, k, options, breakOnAnyOrderedListLine, out var normalizedLazyLine)) {
                    collected.Add(normalizedLazyLine);
                    sourceLines?.Add(new MarkdownSourceLineSlice(normalizedLazyLine, absoluteLineOffset + k + 1, indentColumns + 1));
                    k++;
                    continue;
                }

                break;
            }

            // Strip the required indent; keep the remainder as-is (including additional indentation).
            string cont = StripLeadingIndentColumns(line, continuationIndent);
            int startColumn = continuationIndent + 1;
            startColumn += CountLeadingIndentColumns(cont);
            cont = cont.TrimStart();
            collected.Add(cont);
            sourceLines?.Add(new MarkdownSourceLineSlice(cont, absoluteLineOffset + k + 1, startColumn));
            k++;
        }

        nextIndex = k;
        return collected;
    }

    private static bool TryGetOpenLeadFencedCode(
        IReadOnlyList<string>? collected,
        out string language,
        out char fenceChar,
        out int fenceLength) {
        language = string.Empty;
        fenceChar = '\0';
        fenceLength = 0;

        if (collected == null || collected.Count == 0) {
            return false;
        }

        if (!IsCodeFenceOpen(collected[0] ?? string.Empty, out language, out fenceChar, out fenceLength)) {
            return false;
        }

        for (int i = 1; i < collected.Count; i++) {
            if (IsCodeFenceClose(collected[i] ?? string.Empty, fenceChar, fenceLength)) {
                return false;
            }
        }

        return true;
    }

    private static bool TryNormalizeListLazyContinuationLine(IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options, bool breakOnAnyOrderedListLine, out string normalized) {
        var source = lines != null && index >= 0 && index < lines.Count ? (lines[index] ?? string.Empty) : string.Empty;
        normalized = source;
        if (string.IsNullOrWhiteSpace(source)) return false;

        var trimmed = source.TrimStart();
        if (trimmed.Length == 0) return false;
        if (trimmed.StartsWith(">")) return false;
        if (IsAtxHeading(trimmed, out _, out _)) return false;
        if (LooksLikeHr(trimmed)) return false;
        if (IsCodeFenceOpen(trimmed, out _, out _, out _)) return false;
        if (LooksLikeTableRow(trimmed)) return false;
        if (ShouldTreatAsDefinitionLine(lines, index, options)) return false;
        if (options.Callouts && IsCalloutHeader("> " + trimmed, out _, out _)) return false;
        if (IsUnorderedListLine(trimmed, out _, out _, out _, out _)) return false;
        if (breakOnAnyOrderedListLine ? IsOrderedListLine(trimmed, out _, out _, out _) : IsParagraphInterruptingOrderedListLine(trimmed)) return false;

        if (options.HtmlBlocks && trimmed.StartsWith("<") && !TryParseAngleAutolink(trimmed, 0, out _, out _, out _)) {
            return false;
        }

        normalized = trimmed;
        return true;
    }

    private static bool TryParseNestedFencedCodeBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, out IMarkdownBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.FencedCode) return false;

        string line = lines[index] ?? string.Empty;
        int indent = CountLeadingIndentColumns(line);
        if (indent < continuationIndent) return false;

        string first = StripLeadingIndentColumns(line, continuationIndent);
        int openingFenceIndent = CountLeadingIndentColumns(first);
        if (openingFenceIndent > 3) return false;

        string openingFence = StripLeadingIndentColumns(first, openingFenceIndent);
        if (!IsCodeFenceOpen(openingFence, out string language, out char fenceChar, out int fenceLen)) return false;

        int j = index + 1;
        var code = new StringBuilder();
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            int ind = CountLeadingIndentColumns(raw);
            string sliced = ind >= continuationIndent ? StripLeadingIndentColumns(raw, continuationIndent) : raw.TrimStart();
            if (IsCodeFenceClose(sliced, fenceChar, fenceLen)) { j++; break; }
            int contentIndentToStrip = Math.Min(openingFenceIndent, CountLeadingIndentColumns(sliced));
            if (contentIndentToStrip > 0) {
                sliced = StripLeadingIndentColumns(sliced, contentIndentToStrip);
            }
            code.AppendLine(sliced);
            j++;
        }

        var content = RemoveSingleTrailingLineEnding(code.ToString());
        string? caption = null;
        // Optional caption line (indented like other nested content)
        if (j < lines.Length) {
            var capLine = lines[j] ?? string.Empty;
            if (CountLeadingIndentColumns(capLine) >= continuationIndent) {
                var capSlice = StripLeadingIndentColumns(capLine, continuationIndent);
                if (TryParseCaption(capSlice, out var cap)) { caption = cap; j++; }
            }
        }

        block = CreateParsedFencedBlock(language, content, isFenced: true, caption, options);
        index = j;
        return true;
    }

    private static bool TryParseNestedIndentedCodeBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, out CodeBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.IndentedCodeBlocks) return false;

        string line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) return false;

        int spaces = CountLeadingIndentColumns(line);
        int required = continuationIndent + 4;
        if (spaces < required) return false;

        int j = index;
        var sb = new StringBuilder();
        while (j < lines.Length) {
            string cur = lines[j] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(cur)) {
                int peek = j + 1;
                if (peek >= lines.Length) break;
                int nextSpaces = CountLeadingIndentColumns(lines[peek] ?? string.Empty);
                if (nextSpaces < required) break;
                sb.AppendLine();
                j++;
                continue;
            }

            int curSpaces = CountLeadingIndentColumns(cur);
            if (curSpaces < required) break;
            sb.AppendLine(StripLeadingIndentColumns(cur, required));
            j++;
        }

        block = new CodeBlock(string.Empty, RemoveSingleTrailingLineEnding(sb.ToString()), isFenced: false);
        index = j;
        return true;
    }

    private static bool TryParseNestedQuoteBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, MarkdownReaderState state, out QuoteBlock? quote) {
        quote = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;

        string line = lines[index] ?? string.Empty;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;
        string slice = StripLeadingIndentColumns(line, continuationIndent);
        if (!slice.TrimStart().StartsWith(">")) return false;

        int j = index;
        var collected = new List<string>();
        bool sawQuotedLine = false;
        string? lastQuoteContent = null;
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(raw)) {
                int peek = j + 1;
                if (peek >= lines.Length) break;
                var next = lines[peek] ?? string.Empty;
                if (CountLeadingIndentColumns(next) < continuationIndent) break;
                string nextPart = StripLeadingIndentColumns(next, continuationIndent);
                if (!nextPart.TrimStart().StartsWith(">")) break;
                collected.Add(string.Empty);
                j++;
                continue;
            }

            if (CountLeadingIndentColumns(raw) < continuationIndent) break;
            string part = StripLeadingIndentColumns(raw, continuationIndent);

            if (string.IsNullOrWhiteSpace(part)) {
                int peek = j + 1;
                if (peek >= lines.Length) break;
                var next = lines[peek] ?? string.Empty;
                if (CountLeadingIndentColumns(next) < continuationIndent) break;
                string nextPart = StripLeadingIndentColumns(next, continuationIndent);
                if (!nextPart.TrimStart().StartsWith(">")) break;
                collected.Add(string.Empty);
                j++;
                continue;
            }

            if (part.TrimStart().StartsWith(">")) {
                string quoteContent = StripSingleQuoteMarker(part);
                if (TryNormalizeQuotedListContinuationLine(lastQuoteContent, quoteContent, options, out var normalizedQuotedLine)) {
                    quoteContent = normalizedQuotedLine;
                } else if (TryNormalizeQuotedIndentedParagraphContinuation(lastQuoteContent, quoteContent, options, out var normalizedQuotedParagraphLine)) {
                    quoteContent = normalizedQuotedParagraphLine;
                }

                collected.Add("> " + quoteContent);
                sawQuotedLine = true;
                lastQuoteContent = quoteContent;
                j++;
                continue;
            }

            // Match the top-level quote parser's lazy continuation behavior inside list items too.
            if (!sawQuotedLine) break;
            var previousQuoteContent = lastQuoteContent;
            if (previousQuoteContent == null || previousQuoteContent.Length == 0) break;
            var quoteContext = new[] { previousQuoteContent, part };
            if (!LooksLikeParagraphLine(quoteContext, 0, options) ||
                !TryNormalizeQuoteLazyContinuationLine(quoteContext, 1, options, out var normalizedLazyLine)) break;

            collected.Add("> " + normalizedLazyLine);
            lastQuoteContent = normalizedLazyLine;
            j++;
        }

        if (collected.Count == 0) return false;

        if (TryParseCollectedNestedBlock(collected, options, state, index, out QuoteBlock? parsedQuote)) {
            quote = parsedQuote;
            index = j;
            return true;
        }
        return false;
    }

    private static string StripSingleQuoteMarker(string line) {
        if (string.IsNullOrEmpty(line)) return string.Empty;
        var trimmed = line.TrimStart();
        if (!trimmed.StartsWith(">")) return trimmed;
        return trimmed.Length >= 2 && trimmed[1] == ' ' ? trimmed.Substring(2) : trimmed.Substring(1);
    }

    private static int GetQuoteContentStartColumn(string line) {
        if (string.IsNullOrEmpty(line)) {
            return 1;
        }

        int column = 1;
        int index = 0;
        while (index < line.Length) {
            char ch = line[index];
            if (ch == ' ') {
                column++;
                index++;
                continue;
            }

            if (ch == '\t') {
                column += 4 - ((column - 1) % 4);
                index++;
                continue;
            }

            break;
        }

        if (index < line.Length && line[index] == '>') {
            column++;
            index++;
        }

        if (index < line.Length && line[index] == ' ') {
            column++;
        }

        return column;
    }

    private static bool TryParseNestedTableBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, MarkdownReaderState state, out TableBlock? table) {
        table = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.Tables) return false;

        string line = lines[index] ?? string.Empty;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;
        string slice = StripLeadingIndentColumns(line, continuationIndent);
        if (!LooksLikeTableRow(slice.TrimStart())) return false;

        int j = index;
        var collected = new List<string>();
        while (j < lines.Length) {
            string raw = lines[j] ?? string.Empty;
            if (CountLeadingIndentColumns(raw) < continuationIndent) break;
            string part = StripLeadingIndentColumns(raw, continuationIndent);
            if (string.IsNullOrWhiteSpace(part)) break;
            // Stop when the row no longer looks table-ish.
            if (!LooksLikeTableRow(part.TrimStart()) && !IsAlignmentRow(part.TrimStart())) break;
            collected.Add(part);
            j++;
        }

        if (collected.Count == 0) return false;
        if (TryParseCollectedNestedBlock(collected, options, state, index, out TableBlock? parsedTable)) {
            table = parsedTable;
            index = j;
            return true;
        }
        return false;
    }

    private static bool TryParseCollectedNestedBlock<TBlock>(
        List<string> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        int lineOffset,
        out TBlock? block)
        where TBlock : class, IMarkdownBlock {
        block = null;
        if (lines == null || lines.Count == 0) return false;

        var nested = ParseBlocksFromLines(lines.ToArray(), options, state, lineOffset: lineOffset);
        if (nested.Count == 0 || nested[0] is not TBlock parsedBlock) {
            return false;
        }

        block = parsedBlock;
        return true;
    }

    private static bool TryParseNestedHtmlBlock(string[] lines, ref int index, int continuationIndent, MarkdownReaderOptions options, MarkdownReaderState state, out IMarkdownBlock? block) {
        block = null;
        if (lines == null || index < 0 || index >= lines.Length) return false;
        if (!options.HtmlBlocks) return false;

        string line = lines[index] ?? string.Empty;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;
        string slice = StripLeadingIndentColumns(line, continuationIndent);
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
            if (CountLeadingIndentColumns(raw) < continuationIndent) break;
            collected.Add(StripLeadingIndentColumns(raw, continuationIndent));
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

    private static List<IMarkdownBlock> ParseBlocksFromLines(string[] lines, MarkdownReaderOptions options, MarkdownReaderState state, List<MarkdownSyntaxNode>? syntaxNodes = null, int lineOffset = 0) {
        var doc = MarkdownDoc.Create();
        var opt = CloneOptionsWithoutFrontMatter(options);
        var pipeline = MarkdownReaderPipeline.Default(opt);
        int previousLineOffset = state.SourceLineOffset;
        state.SourceLineOffset = lineOffset;

        try {
            int i = 0;
            while (i < lines.Length) {
                if (string.IsNullOrWhiteSpace(lines[i])) { i++; continue; }
                bool matched = false;
                var parsers = pipeline.Parsers;
                int previousBlockCount = doc.Blocks.Count;
                int startLine = lineOffset + i;
                for (int p = 0; p < parsers.Count; p++) {
                    if (parsers[p].TryParse(lines, ref i, opt, doc, state)) {
                        matched = true;
                        if (syntaxNodes != null && doc.Blocks.Count > previousBlockCount) {
                            CaptureSyntaxNodes(doc, previousBlockCount, startLine, lineOffset + i, syntaxNodes, state);
                        }
                        break;
                    }
                }
                if (!matched) i++;
            }
        } finally {
            state.SourceLineOffset = previousLineOffset;
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

    private readonly struct MarkdownSourceLineSlice {
        public MarkdownSourceLineSlice(string text, int absoluteLine, int startColumn) {
            Text = text ?? string.Empty;
            AbsoluteLine = absoluteLine;
            StartColumn = startColumn < 1 ? 1 : startColumn;
        }

        public string Text { get; }
        public int AbsoluteLine { get; }
        public int StartColumn { get; }
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

    private static List<InlineSequence> ParseParagraphsFromSourceLines(List<MarkdownSourceLineSlice> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        var paragraphs = new List<InlineSequence>();
        if (lines == null || lines.Count == 0) {
            paragraphs.Add(ParseInlines(string.Empty, options, state));
            return paragraphs;
        }

        var current = new List<MarkdownSourceLineSlice>();
        for (int i = 0; i < lines.Count; i++) {
            if (string.IsNullOrEmpty(lines[i].Text)) {
                if (current.Count > 0) {
                    var (text, sourceMap) = JoinParagraphSourceLinesWithSourceMap(current, options, state);
                    paragraphs.Add(ParseInlines(text, options, state, sourceMap));
                    current.Clear();
                }
                continue;
            }

            current.Add(lines[i]);
        }

        if (current.Count > 0) {
            var (text, sourceMap) = JoinParagraphSourceLinesWithSourceMap(current, options, state);
            paragraphs.Add(ParseInlines(text, options, state, sourceMap));
        }

        if (paragraphs.Count == 0) {
            paragraphs.Add(ParseInlines(string.Empty, options, state));
        }

        return paragraphs;
    }

    private static List<ParagraphBlock> ParseParagraphBlocksFromLines(List<string> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        var paragraphInlines = ParseParagraphsFromLines(lines, options, state);
        var blocks = new List<ParagraphBlock>(paragraphInlines.Count);
        for (int i = 0; i < paragraphInlines.Count; i++) {
            blocks.Add(new ParagraphBlock(paragraphInlines[i]));
        }
        return blocks;
    }

    private static List<ParagraphBlock> ParseParagraphBlocksFromSourceLines(List<MarkdownSourceLineSlice> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        var paragraphInlines = ParseParagraphsFromSourceLines(lines, options, state);
        var blocks = new List<ParagraphBlock>(paragraphInlines.Count);
        for (int i = 0; i < paragraphInlines.Count; i++) {
            blocks.Add(new ParagraphBlock(paragraphInlines[i]));
        }
        return blocks;
    }

    private static void AddListItemLeadSyntaxNodes(ListItem item, List<string> lines, int lineOffset, MarkdownReaderOptions options, MarkdownReaderState? state, List<MarkdownSourceLineSlice>? sourceLines = null) {
        if (item == null || lines == null || lines.Count == 0) return;
        if (item.SyntaxChildren.Count > 0) return;
        int absoluteLineOffset = (state?.SourceLineOffset ?? 0) + lineOffset;

        if (TryParseListItemLeadBlockSyntaxNodes(lines, lineOffset, options, state, sourceLines, out var leadBlockSyntax)) {
            for (int i = 0; i < leadBlockSyntax.Count; i++) {
                item.SyntaxChildren.Add(leadBlockSyntax[i]);
            }
            return;
        }

        if (TryParseSetextHeadingParagraphLines(lines, options, out int level, out string headingText)) {
            item.SyntaxChildren.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.Heading,
                CreateLineSpan(state, absoluteLineOffset + 1, absoluteLineOffset + lines.Count),
                headingText));
            return;
        }

        if (TryGetLeadingSetextHeadingPrefix(lines, options, out int headingLineCount, out level, out headingText)) {
            item.SyntaxChildren.Add(new MarkdownSyntaxNode(
                MarkdownSyntaxKind.Heading,
                CreateLineSpan(state, absoluteLineOffset + 1, absoluteLineOffset + headingLineCount),
                headingText));

            if (headingLineCount < lines.Count) {
                var trailingLines = lines.GetRange(headingLineCount, lines.Count - headingLineCount);
                if (!trailingLines.TrueForAll(string.IsNullOrWhiteSpace)) {
                    IReadOnlyList<MarkdownSyntaxNode> trailingSyntax;
                    if (sourceLines != null && sourceLines.Count >= headingLineCount) {
                        trailingSyntax = ParseBlockSyntaxNodesFromSourceLines(sourceLines.GetRange(headingLineCount, lines.Count - headingLineCount), options, state);
                    } else {
                        var nestedSyntax = new List<MarkdownSyntaxNode>();
                        ParseBlocksFromLines(trailingLines.ToArray(), options, state ?? new MarkdownReaderState(), nestedSyntax, lineOffset: lineOffset + headingLineCount);
                        trailingSyntax = nestedSyntax;
                    }

                    for (int i = 0; i < trailingSyntax.Count; i++) {
                        item.SyntaxChildren.Add(trailingSyntax[i]);
                    }
                }
            }
            return;
        }

        int firstBlank = lines.FindIndex(string.IsNullOrWhiteSpace);
        if (firstBlank > 0) {
            if (sourceLines != null && sourceLines.Count >= firstBlank) {
                AddParagraphSyntaxNodes(item.SyntaxChildren, sourceLines.GetRange(0, firstBlank), options, state);
            } else {
                AddParagraphSyntaxNodes(item.SyntaxChildren, lines.GetRange(0, firstBlank), absoluteLineOffset, options, state);
            }

            if (firstBlank + 1 < lines.Count) {
                var trailingLines = lines.GetRange(firstBlank + 1, lines.Count - firstBlank - 1);
                if (!trailingLines.TrueForAll(string.IsNullOrWhiteSpace)) {
                    IReadOnlyList<MarkdownSyntaxNode> trailingSyntax;
                    if (sourceLines != null && sourceLines.Count > firstBlank + 1) {
                        trailingSyntax = ParseBlockSyntaxNodesFromSourceLines(sourceLines.GetRange(firstBlank + 1, lines.Count - firstBlank - 1), options, state);
                    } else {
                        var nestedSyntax = new List<MarkdownSyntaxNode>();
                        ParseBlocksFromLines(trailingLines.ToArray(), options, state ?? new MarkdownReaderState(), nestedSyntax, lineOffset: lineOffset + firstBlank + 1);
                        trailingSyntax = nestedSyntax;
                    }

                    for (int i = 0; i < trailingSyntax.Count; i++) {
                        item.SyntaxChildren.Add(trailingSyntax[i]);
                    }
                    return;
                }
            }

            return;
        }

        if (sourceLines != null && sourceLines.Count == lines.Count) {
            AddParagraphSyntaxNodes(item.SyntaxChildren, sourceLines, options, state);
        } else {
            AddParagraphSyntaxNodes(item.SyntaxChildren, lines, absoluteLineOffset, options, state);
        }
    }

    private static void AddParagraphSyntaxNodes(List<MarkdownSyntaxNode> nodes, List<string> lines, int lineOffset, MarkdownReaderOptions options, MarkdownReaderState? state) {
        if (nodes == null || lines == null || lines.Count == 0) return;

        var current = new List<string>();
        int currentStart = -1;

        for (int i = 0; i < lines.Count; i++) {
            var line = lines[i] ?? string.Empty;
            if (line.Length == 0) {
                FlushParagraphSyntaxNode(nodes, current, currentStart, i - 1, lineOffset, options, state);
                current.Clear();
                currentStart = -1;
                continue;
            }

            if (currentStart < 0) currentStart = i;
            current.Add(line);
        }

        FlushParagraphSyntaxNode(nodes, current, currentStart, lines.Count - 1, lineOffset, options, state);
    }

    private static void AddParagraphSyntaxNodes(List<MarkdownSyntaxNode> nodes, List<MarkdownSourceLineSlice> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        if (nodes == null || lines == null || lines.Count == 0) return;

        var current = new List<MarkdownSourceLineSlice>();
        for (int i = 0; i < lines.Count; i++) {
            if (string.IsNullOrEmpty(lines[i].Text)) {
                FlushParagraphSyntaxNode(nodes, current, options, state);
                current.Clear();
                continue;
            }

            current.Add(lines[i]);
        }

        FlushParagraphSyntaxNode(nodes, current, options, state);
    }

    private static IReadOnlyList<MarkdownSyntaxNode> ParseBlockSyntaxNodesFromSourceLines(
        List<MarkdownSourceLineSlice> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState? state) {
        if (lines == null || lines.Count == 0) {
            return Array.Empty<MarkdownSyntaxNode>();
        }

        var effectiveState = state ?? new MarkdownReaderState();
        var (_, syntaxChildren) = ParseNestedMarkdownBlocks(lines, options, effectiveState);
        return syntaxChildren;
    }

    private static void FlushParagraphSyntaxNode(List<MarkdownSyntaxNode> nodes, List<string> lines, int startIndex, int endIndex, int lineOffset, MarkdownReaderOptions options, MarkdownReaderState? state) {
        if (nodes == null || lines == null || lines.Count == 0 || startIndex < 0 || endIndex < startIndex) return;

        var inlines = ParseInlines(JoinParagraphLines(lines, options), options, state);
        var paragraph = new ParagraphBlock(inlines);
        nodes.Add(BuildSyntaxNode(paragraph, CreateLineSpan(state, lineOffset + startIndex + 1, lineOffset + endIndex + 1)));
    }

    private static void FlushParagraphSyntaxNode(List<MarkdownSyntaxNode> nodes, List<MarkdownSourceLineSlice> lines, MarkdownReaderOptions options, MarkdownReaderState? state) {
        if (nodes == null || lines == null || lines.Count == 0) return;

        var (text, sourceMap) = JoinParagraphSourceLinesWithSourceMap(lines, options, state);
        var inlines = ParseInlines(text, options, state, sourceMap);
        var paragraph = new ParagraphBlock(inlines);
        nodes.Add(BuildSyntaxNode(paragraph, CreateSpan(
            state,
            lines[0].AbsoluteLine,
            lines[0].StartColumn,
            lines[lines.Count - 1].AbsoluteLine,
            lines[lines.Count - 1].StartColumn + Math.Max(0, lines[lines.Count - 1].Text.Length - 1))));
    }

    private static void AddListItemChildSyntaxNode(ListItem item, IMarkdownBlock block, int startLineIndex, int endExclusiveLineIndex, MarkdownReaderState? state) {
        if (item == null || block == null) return;
        int absoluteStart = (state?.SourceLineOffset ?? 0) + startLineIndex;
        int absoluteEndExclusive = (state?.SourceLineOffset ?? 0) + endExclusiveLineIndex;
        item.SyntaxChildren.Add(BuildSyntaxNode(block, CreateLineSpan(state, absoluteStart + 1, Math.Max(absoluteStart + 1, absoluteEndExclusive))));
    }

    private static ListItem CreateListItemFromLeadLines(List<string> lines, bool isTask, bool done, MarkdownReaderOptions options, MarkdownReaderState? state, List<MarkdownSourceLineSlice>? sourceLines = null) {
        if (TryCreateListItemFromLeadBlocks(lines, isTask, done, options, state, sourceLines, out var blockLeadItem)) {
            return blockLeadItem;
        }

        if (TryParseListItemLeadSetextBlocks(lines, options, state, out var leadBlocks)) {
            var headingItem = isTask ? ListItem.TaskInlines(new InlineSequence(), done) : new ListItem(new InlineSequence());
            for (int i = 0; i < leadBlocks.Count; i++) {
                headingItem.Children.Add(leadBlocks[i]);
            }
            return headingItem;
        }

        int firstBlank = lines.FindIndex(string.IsNullOrWhiteSpace);
        if (firstBlank <= 0) {
            var paragraphs = sourceLines != null && sourceLines.Count == lines.Count
                ? ParseParagraphsFromSourceLines(sourceLines, options, state)
                : ParseParagraphsFromLines(lines, options, state);
            var item = isTask ? ListItem.TaskInlines(paragraphs[0], done) : new ListItem(paragraphs[0]);
            for (int i = 1; i < paragraphs.Count; i++) {
                item.AdditionalParagraphs.Add(paragraphs[i]);
            }
            return item;
        }

        var firstParagraph = sourceLines != null && sourceLines.Count >= firstBlank
            ? ParseParagraphsFromSourceLines(sourceLines.GetRange(0, firstBlank), options, state)[0]
            : ParseParagraphsFromLines(lines.GetRange(0, firstBlank), options, state)[0];
        var mixedItem = isTask ? ListItem.TaskInlines(firstParagraph, done) : new ListItem(firstParagraph);

        if (firstBlank + 1 >= lines.Count) return mixedItem;

        var trailingLines = lines.GetRange(firstBlank + 1, lines.Count - firstBlank - 1);
        if (trailingLines.TrueForAll(string.IsNullOrWhiteSpace)) return mixedItem;

        if (sourceLines != null && sourceLines.Count > firstBlank + 1) {
            var trailingSourceLines = sourceLines.GetRange(firstBlank + 1, lines.Count - firstBlank - 1);
            if (!trailingSourceLines.TrueForAll(slice => string.IsNullOrWhiteSpace(slice.Text))) {
                var effectiveState = state ?? new MarkdownReaderState();
                var (trailingBlocksFromSource, trailingSyntaxFromSource) = ParseNestedMarkdownBlocks(trailingSourceLines, options, effectiveState);
                if (trailingSyntaxFromSource.All(node => node.Kind == MarkdownSyntaxKind.Paragraph)) {
                    var trailingParagraphs = ParseParagraphsFromSourceLines(trailingSourceLines, options, state);
                    for (int i = 0; i < trailingParagraphs.Count; i++) {
                        mixedItem.AdditionalParagraphs.Add(trailingParagraphs[i]);
                    }
                    return mixedItem;
                }

                for (int i = 0; i < trailingBlocksFromSource.Count; i++) {
                    mixedItem.Children.Add(trailingBlocksFromSource[i]);
                }
                mixedItem.ForceLoose = true;
                return mixedItem;
            }
        }

        var trailingBlocks = ParseBlocksFromLines(trailingLines.ToArray(), options, state ?? new MarkdownReaderState());
        if (mixedItem.TryAbsorbTrailingParagraphBlocks(trailingBlocks)) return mixedItem;

        for (int i = 0; i < trailingBlocks.Count; i++) {
            mixedItem.Children.Add(trailingBlocks[i]);
        }
        mixedItem.ForceLoose = true;
        return mixedItem;
    }

    private static bool TryCreateListItemFromLeadBlocks(
        List<string> lines,
        bool isTask,
        bool done,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        List<MarkdownSourceLineSlice>? sourceLines,
        out ListItem item) {
        item = null!;
        IReadOnlyList<IMarkdownBlock> leadBlocks;
        IReadOnlyList<MarkdownSyntaxNode> leadSyntax = Array.Empty<MarkdownSyntaxNode>();

        if (lines == null || lines.Count == 0 || !StartsListItemLeadWithStandaloneBlock(lines, options)) {
            return false;
        }

        if (sourceLines != null && sourceLines.Count == lines.Count) {
            var (blocksFromSource, syntaxFromSource) = ParseNestedMarkdownBlocks(sourceLines, options, state ?? new MarkdownReaderState());
            if (blocksFromSource.Count == 0 || blocksFromSource.All(block => block is ParagraphBlock)) {
                return false;
            }

            leadBlocks = blocksFromSource;
            leadSyntax = syntaxFromSource;
        } else if (!TryParseListItemLeadBlocks(lines, options, state, sourceLines, out leadBlocks)) {
            return false;
        }

        var blockLeadItem = isTask ? ListItem.TaskInlines(new InlineSequence(), done) : new ListItem(new InlineSequence());
        for (int i = 0; i < leadBlocks.Count; i++) {
            blockLeadItem.Children.Add(leadBlocks[i]);
        }
        for (int i = 0; i < leadSyntax.Count; i++) {
            blockLeadItem.SyntaxChildren.Add(leadSyntax[i]);
        }
        if (leadBlocks.Count > 1 && lines.Exists(string.IsNullOrWhiteSpace)) {
            blockLeadItem.ForceLoose = true;
        }

        item = blockLeadItem;
        return true;
    }

    private static bool TryParseListItemLeadBlocks(
        List<string> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        List<MarkdownSourceLineSlice>? sourceLines,
        out IReadOnlyList<IMarkdownBlock> blocks) {
        blocks = Array.Empty<IMarkdownBlock>();
        if (lines == null || lines.Count == 0) {
            return false;
        }

        if (!StartsListItemLeadWithStandaloneBlock(lines, options)) {
            return false;
        }

        IReadOnlyList<IMarkdownBlock> parsedBlocks;
        if (sourceLines != null && sourceLines.Count == lines.Count) {
            var (blocksFromSource, _) = ParseNestedMarkdownBlocks(sourceLines, options, state ?? new MarkdownReaderState());
            parsedBlocks = blocksFromSource;
        } else {
            parsedBlocks = ParseBlocksFromLines(lines.ToArray(), options, state ?? new MarkdownReaderState());
        }

        if (parsedBlocks.Count == 0 || parsedBlocks.All(block => block is ParagraphBlock)) {
            return false;
        }

        blocks = parsedBlocks;
        return true;
    }

    private static bool TryParseListItemLeadBlockSyntaxNodes(
        List<string> lines,
        int lineOffset,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        List<MarkdownSourceLineSlice>? sourceLines,
        out IReadOnlyList<MarkdownSyntaxNode> syntaxNodes) {
        syntaxNodes = Array.Empty<MarkdownSyntaxNode>();
        if (lines == null || lines.Count == 0) {
            return false;
        }

        if (!StartsListItemLeadWithStandaloneBlock(lines, options)) {
            return false;
        }

        IReadOnlyList<MarkdownSyntaxNode> parsedSyntax;
        if (sourceLines != null && sourceLines.Count == lines.Count) {
            parsedSyntax = ParseBlockSyntaxNodesFromSourceLines(sourceLines, options, state);
        } else {
            var leadSyntax = new List<MarkdownSyntaxNode>();
            ParseBlocksFromLines(lines.ToArray(), options, state ?? new MarkdownReaderState(), leadSyntax, lineOffset);
            parsedSyntax = leadSyntax;
        }

        if (parsedSyntax.Count == 0 || parsedSyntax.All(node => node.Kind == MarkdownSyntaxKind.Paragraph)) {
            return false;
        }

        syntaxNodes = parsedSyntax;
        return true;
    }

    private static bool StartsListItemLeadWithIndentedCode(List<string> lines, MarkdownReaderOptions options) {
        if (lines == null || lines.Count == 0 || options?.IndentedCodeBlocks != true) {
            return false;
        }

        int firstNonBlank = lines.FindIndex(line => !string.IsNullOrWhiteSpace(line));
        if (firstNonBlank < 0) {
            return false;
        }

        return CountLeadingIndentColumns(lines[firstNonBlank] ?? string.Empty) >= 4;
    }

    private static bool StartsListItemLeadWithStandaloneBlock(List<string> lines, MarkdownReaderOptions options) {
        if (lines == null || lines.Count == 0 || options == null) {
            return false;
        }

        if (StartsListItemLeadWithIndentedCode(lines, options)) {
            return true;
        }

        int firstNonBlank = lines.FindIndex(line => !string.IsNullOrWhiteSpace(line));
        if (firstNonBlank < 0) {
            return false;
        }

        var firstLine = lines[firstNonBlank] ?? string.Empty;
        var trimmed = firstLine.TrimStart();
        if (trimmed.Length == 0) {
            return false;
        }

        if (options.Headings && trimmed[0] == '#') {
            int markerLength = 0;
            while (markerLength < trimmed.Length && trimmed[markerLength] == '#') {
                markerLength++;
            }

            if (markerLength > 0 && markerLength <= 6 && (markerLength == trimmed.Length || char.IsWhiteSpace(trimmed[markerLength]))) {
                return true;
            }
        }

        if (trimmed[0] == '>') {
            return true;
        }

        if (options.FencedCode && IsCodeFenceOpen(trimmed, out _, out _, out _)) {
            return true;
        }

        if (options.UnorderedLists && IsUnorderedListLine(trimmed, out _, out _, out _)) {
            return true;
        }

        if (options.OrderedLists && IsOrderedListLine(trimmed, out _, out _)) {
            return true;
        }

        return false;
    }

    private static bool TryParseListItemLeadSetextBlocks(List<string> lines, MarkdownReaderOptions options, MarkdownReaderState? state, out List<IMarkdownBlock> blocks) {
        blocks = new List<IMarkdownBlock>();
        if (lines == null || lines.Count == 0 || options == null || !options.Headings) return false;

        if (!TryGetLeadingSetextHeadingPrefix(lines, options, out int headingLineCount, out int level, out string headingText)) return false;

        blocks.Add(new HeadingBlock(level, ParseInlines(headingText, options, state)));

        if (headingLineCount >= lines.Count) return true;

        var trailingLines = lines.GetRange(headingLineCount, lines.Count - headingLineCount);
        if (trailingLines.TrueForAll(string.IsNullOrWhiteSpace)) return true;

        var trailingBlocks = ParseBlocksFromLines(trailingLines.ToArray(), options, state ?? new MarkdownReaderState());
        for (int i = 0; i < trailingBlocks.Count; i++) {
            blocks.Add(trailingBlocks[i]);
        }

        return true;
    }

    private static bool TryGetLeadingSetextHeadingPrefix(List<string> lines, MarkdownReaderOptions options, out int headingLineCount, out int level, out string headingText) {
        headingLineCount = 0;
        level = 0;
        headingText = string.Empty;
        if (lines == null || lines.Count < 2 || options == null || !options.Headings) return false;

        int firstBlank = lines.FindIndex(string.IsNullOrWhiteSpace);
        int maxPrefixLength = firstBlank >= 0 ? firstBlank : lines.Count;
        if (maxPrefixLength < 2) return false;

        for (int prefixLength = 2; prefixLength <= maxPrefixLength; prefixLength++) {
            var candidate = lines.GetRange(0, prefixLength);
            if (!TryParseSetextHeadingParagraphLines(candidate, options, out level, out headingText)) continue;

            headingLineCount = prefixLength;
            return true;
        }

        level = 0;
        headingText = string.Empty;
        return false;
    }

    private static bool TryParseSetextHeadingParagraphLines(List<string> lines, MarkdownReaderOptions options, out int level, out string headingText) {
        level = 0;
        headingText = string.Empty;
        if (lines == null || lines.Count < 2 || options == null || !options.Headings) return false;

        var underline = lines[lines.Count - 1]?.Trim() ?? string.Empty;
        if (underline.Length < 3) return false;

        char marker = '\0';
        for (int i = 0; i < underline.Length; i++) {
            char ch = underline[i];
            if (ch != '=' && ch != '-') return false;
            if (marker == '\0') marker = ch;
            else if (ch != marker) return false;
        }

        var contentLines = lines.GetRange(0, lines.Count - 1);
        if (contentLines.Count == 0 || contentLines.TrueForAll(string.IsNullOrWhiteSpace)) return false;

        level = marker == '=' ? 1 : 2;
        headingText = JoinParagraphLines(contentLines, options).Trim();
        return headingText.Length > 0;
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

    private static (string Text, MarkdownInlineSourceMap? SourceMap) JoinParagraphLinesWithSourceMap(
        List<string> lines,
        int absoluteLineOffset,
        MarkdownReaderOptions options,
        MarkdownReaderState? state) {
        var text = JoinParagraphLines(lines, options);
        if (state?.SourceTextMap == null || string.IsNullOrEmpty(text)) {
            return (text, null);
        }

        var points = new MarkdownSourcePoint?[text.Length];
        var cursor = 0;
        var previousLineForJoin = absoluteLineOffset + 1;
        var previousJoinColumn = 1;

        for (var i = 0; i < lines.Count; i++) {
            var raw = lines[i] ?? string.Empty;
            var trimmed = raw.TrimEnd();
            trimmed = ConsumeTrailingBackslashHardBreak(trimmed, options, out _);

            if (i > 0 && cursor < points.Length) {
                points[cursor++] = state.SourceTextMap.CreatePoint(previousLineForJoin, previousJoinColumn);
            }

            var absoluteLine = absoluteLineOffset + i + 1;
            for (var charIndex = 0; charIndex < trimmed.Length && cursor < points.Length; charIndex++) {
                points[cursor++] = state.SourceTextMap.CreatePoint(absoluteLine, charIndex + 1);
            }

            previousLineForJoin = absoluteLine;
            previousJoinColumn = Math.Max(1, trimmed.Length);
        }

        if (cursor < points.Length) {
            Array.Resize(ref points, cursor);
        }

        return (text, new MarkdownInlineSourceMap(points));
    }

    private static (string Text, MarkdownInlineSourceMap? SourceMap) JoinParagraphSourceLinesWithSourceMap(
        List<MarkdownSourceLineSlice> lines,
        MarkdownReaderOptions options,
        MarkdownReaderState? state) {
        if (lines == null || lines.Count == 0) {
            return (string.Empty, null);
        }

        var plainLines = new List<string>(lines.Count);
        for (int i = 0; i < lines.Count; i++) {
            plainLines.Add(lines[i].Text);
        }

        var text = JoinParagraphLines(plainLines, options);
        if (state?.SourceTextMap == null || string.IsNullOrEmpty(text)) {
            return (text, null);
        }

        var points = new MarkdownSourcePoint?[text.Length];
        var cursor = 0;
        var previousLine = lines[0].AbsoluteLine;
        var previousJoinColumn = lines[0].StartColumn;

        for (var i = 0; i < lines.Count; i++) {
            if (i > 0 && cursor < points.Length) {
                points[cursor++] = state.SourceTextMap.CreatePoint(previousLine, previousJoinColumn);
            }

            var slice = lines[i];
            for (var charIndex = 0; charIndex < slice.Text.Length && cursor < points.Length; charIndex++) {
                points[cursor++] = state.SourceTextMap.CreatePoint(slice.AbsoluteLine, slice.StartColumn + charIndex);
            }

            previousLine = slice.AbsoluteLine;
            previousJoinColumn = slice.StartColumn + Math.Max(0, slice.Text.Length - 1);
        }

        if (cursor < points.Length) {
            Array.Resize(ref points, cursor);
        }

        return (text, new MarkdownInlineSourceMap(points));
    }

    private static MarkdownInlineSourceMap? BuildInlineSourceMapForSingleLine(
        string text,
        int absoluteLine,
        int startColumn,
        MarkdownReaderState? state) {
        if (state?.SourceTextMap == null || string.IsNullOrEmpty(text)) {
            return null;
        }

        var points = new MarkdownSourcePoint?[text.Length];
        for (var i = 0; i < text.Length; i++) {
            points[i] = state.SourceTextMap.CreatePoint(absoluteLine, startColumn + i);
        }

        return new MarkdownInlineSourceMap(points);
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
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        ListItem item,
        bool allowNestedOrdered,
        bool allowNestedUnordered) {

        if (lines == null || item == null) return;

        while (index < lines.Length) {
            if (IsStructurallyBlankListItem(item) && string.IsNullOrWhiteSpace(lines[index])) {
                return;
            }

            int k = index;
            bool sawBlankLine = false;

            // Skip blank lines only when they are followed by nested content.
            while (k < lines.Length && string.IsNullOrWhiteSpace(lines[k])) {
                sawBlankLine = true;
                int peek = k + 1;
                if (peek >= lines.Length) return;
                var next = lines[peek] ?? string.Empty;
                if (string.IsNullOrWhiteSpace(next)) {
                    k = peek;
                    continue;
                }
                if (CountLeadingIndentColumns(next) < continuationIndent) return;
                if (!IsListNestedBlockStart(next, continuationIndent, itemLevelAbs, allowNestedOrdered, allowNestedUnordered, options)) {
                    k = peek;
                    break;
                }
                k = peek;
            }

            if (k >= lines.Length) { index = k; return; }
            if (!sawBlankLine && k > 0 && string.IsNullOrWhiteSpace(lines[k - 1])) sawBlankLine = true;

            // Nested fenced code block
            int tmp = k;
            if (TryParseNestedFencedCodeBlock(lines, ref tmp, continuationIndent, options, out var code) && code != null) {
                item.Children.Add(code);
                AddListItemChildSyntaxNode(item, code, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested indented code block
            tmp = k;
            if (TryParseNestedIndentedCodeBlock(lines, ref tmp, continuationIndent, options, out var indented) && indented != null) {
                item.Children.Add(indented);
                AddListItemChildSyntaxNode(item, indented, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested blockquote
            tmp = k;
            if (TryParseNestedQuoteBlock(lines, ref tmp, continuationIndent, options, state, out var quote) && quote != null) {
                item.Children.Add(quote);
                AddListItemChildSyntaxNode(item, quote, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested table
            tmp = k;
            if (TryParseNestedTableBlock(lines, ref tmp, continuationIndent, options, state, out var table) && table != null) {
                item.Children.Add(table);
                AddListItemChildSyntaxNode(item, table, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested HTML blocks (details / raw HTML) when HtmlBlocks are enabled.
            tmp = k;
            if (TryParseNestedHtmlBlock(lines, ref tmp, continuationIndent, options, state, out var htmlBlock) && htmlBlock != null) {
                item.Children.Add(htmlBlock);
                AddListItemChildSyntaxNode(item, htmlBlock, k, tmp, state);
                if (sawBlankLine) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nested ordered list
            if (allowNestedOrdered
                && options.OrderedLists
                && CountLeadingIndentColumns(lines[k] ?? string.Empty) >= continuationIndent
                && IsOrderedListLine(lines[k], out int lvlAbsO2, out _, out _)
                && lvlAbsO2 >= itemLevelAbs + 1) {
                if (TryParseNestedListBlock(lines, k, options, state, new OrderedListParser(), out var orderedList, out var orderedEndIndex)) {
                    item.Children.Add(orderedList);
                    AddListItemChildSyntaxNode(item, orderedList, k, orderedEndIndex, state);
                    if (sawBlankLine) item.ForceLoose = true;
                    index = orderedEndIndex;
                    continue;
                }
            }

            // Nested unordered list
            if (allowNestedUnordered
                && options.UnorderedLists
                && CountLeadingIndentColumns(lines[k] ?? string.Empty) >= continuationIndent
                && IsUnorderedListLine(lines[k], out int lvlAbsU2, out _, out _, out _)
                && lvlAbsU2 >= itemLevelAbs + 1) {
                if (TryParseNestedListBlock(lines, k, options, state, new UnorderedListParser(), out var unorderedList, out var unorderedEndIndex)) {
                    item.Children.Add(unorderedList);
                    AddListItemChildSyntaxNode(item, unorderedList, k, unorderedEndIndex, state);
                    if (sawBlankLine) item.ForceLoose = true;
                    index = unorderedEndIndex;
                    continue;
                }
            }

            tmp = k;
            if (TryParseTrailingParagraphsForListItem(lines, ref tmp, itemLevelAbs, continuationIndent, options, state, out var trailingParagraphs, out var trailingSyntaxNodes) && trailingParagraphs.Count > 0) {
                foreach (var paragraph in trailingParagraphs) {
                    item.Children.Add(paragraph);
                }
                for (int p = 0; p < trailingSyntaxNodes.Count; p++) {
                    item.SyntaxChildren.Add(trailingSyntaxNodes[p]);
                }
                if (sawBlankLine || item.Children.Count > 0) item.ForceLoose = true;
                index = tmp;
                continue;
            }

            // Nothing nested to consume.
            index = k;
            return;
        }
    }

    private static bool IsStructurallyBlankListItem(ListItem item) {
        return item.Content.Nodes.Count == 0
               && item.AdditionalParagraphs.Count == 0
               && item.Children.Count == 0;
    }

    private static bool TryParseNestedListBlock(
        string[] lines,
        int startIndex,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        IMarkdownBlockParser parser,
        out IMarkdownListBlock list,
        out int endIndex) {
        int idx = startIndex;
        var tempDoc = MarkdownDoc.Create();
        var effectiveState = state ?? new MarkdownReaderState();
        if (parser.TryParse(lines, ref idx, options, tempDoc, effectiveState) &&
            tempDoc.Blocks.Count == 1 &&
            tempDoc.Blocks[0] is IMarkdownListBlock parsedList) {
            list = parsedList;
            endIndex = idx;
            return true;
        }

        list = null!;
        endIndex = startIndex;
        return false;
    }

    private static bool TryParseTrailingParagraphsForListItem(
        string[] lines,
        ref int index,
        int itemLevelAbs,
        int continuationIndent,
        MarkdownReaderOptions options,
        MarkdownReaderState state,
        out List<ParagraphBlock> paragraphs,
        out List<MarkdownSyntaxNode> syntaxNodes) {

        paragraphs = new List<ParagraphBlock>();
        syntaxNodes = new List<MarkdownSyntaxNode>();
        if (lines == null || index < 0 || index >= lines.Length) return false;

        string line = lines[index] ?? string.Empty;
        if (string.IsNullOrWhiteSpace(line)) return false;
        if (CountLeadingIndentColumns(line) < continuationIndent) return false;
        if (IsListNestedBlockStart(line, continuationIndent, itemLevelAbs, allowNestedOrdered: true, allowNestedUnordered: true, options)) return false;
        if (IsUnorderedListLine(line, out _, out _, out _, out _) || IsOrderedListLine(line, out _, out _, out _)) return false;

        string firstContent = StripLeadingIndentColumns(line, continuationIndent);
        firstContent = firstContent.TrimStart();
        int firstStartColumn = continuationIndent + CountLeadingIndentColumns(StripLeadingIndentColumns(line, continuationIndent)) + 1;

        int next = index + 1;
        var paragraphSourceLines = new List<MarkdownSourceLineSlice>();
        var paragraphLines = ConsumeListContinuationLines(
            lines,
            ref next,
            continuationIndent,
            firstContent,
            options,
            sourceLines: paragraphSourceLines,
            absoluteLineOffset: state.SourceLineOffset,
            initialLineIndex: index,
            initialStartColumn: firstStartColumn);
        paragraphs.AddRange(ParseParagraphBlocksFromSourceLines(paragraphSourceLines, options, state));
        AddParagraphSyntaxNodes(syntaxNodes, paragraphSourceLines, options, state);

        index = next;
        return paragraphs.Count > 0;
    }

    private static bool IsListNestedBlockStart(
        string line,
        int continuationIndent,
        int itemLevelAbs,
        bool allowNestedOrdered,
        bool allowNestedUnordered,
        MarkdownReaderOptions options) {

        if (string.IsNullOrEmpty(line)) return false;

        int nextIndentColumns = CountLeadingIndentColumns(line);
        if (nextIndentColumns < continuationIndent) return false;

        if (allowNestedOrdered && options.OrderedLists &&
            IsOrderedListLine(line, out int lvlAbsO, out _, out _) &&
            lvlAbsO >= itemLevelAbs + 1) {
            return true;
        }

        if (allowNestedUnordered && options.UnorderedLists &&
            IsUnorderedListLine(line, out int lvlAbsU, out _, out _, out _) &&
            lvlAbsU >= itemLevelAbs + 1) {
            return true;
        }

        var slice = StripLeadingIndentColumns(line, continuationIndent);
        var sliceTrim = slice.TrimStart();

        if (options.FencedCode && IsCodeFenceOpen(slice, out _, out _, out _)) return true;
        if (options.IndentedCodeBlocks && nextIndentColumns >= continuationIndent + 4 && !string.IsNullOrWhiteSpace(slice)) return true;
        if (sliceTrim.StartsWith(">")) return true;

        if (options.Tables && LooksLikeTableRow(sliceTrim)) return true;

        if (options.HtmlBlocks && sliceTrim.StartsWith("<") && !TryParseAngleAutolink(sliceTrim, 0, out _, out _, out _)) {
            return true;
        }

        return false;
    }

    private static bool IsDefinitionLine(string line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        var trimmed = line.TrimStart();
        if (IsAtxHeading(trimmed, out _, out _)) return false; // headings take priority over definition lists
        if (IsUnorderedListLine(trimmed, out _, out _, out _)) return false; // list items with ":" are not definition terms
        if (IsOrderedListLine(trimmed, out _, out _)) return false; // numbered list items with ":" are not definition terms
        if (StartsWithReferenceDefinitionLikeLabel(trimmed)) return false; // malformed or valid link ref definitions should not become <dl>
        return TryGetDefinitionSeparator(line, out _);
    }

    private static bool ShouldTreatAsDefinitionLine(IReadOnlyList<string>? lines, int index, MarkdownReaderOptions options) {
        if (lines == null || index < 0 || index >= lines.Count) return false;
        if (options == null || !options.DefinitionLists) return false;

        var line = lines[index] ?? string.Empty;
        if (!IsDefinitionLineBlockCandidate(line)) return false;
        if (!options.PreferNarrativeSingleLineDefinitions) return true;

        return HasAdjacentDefinitionLine(lines, index) || HasDefinitionContinuation(lines, index);
    }

    private static bool HasAdjacentDefinitionLine(IReadOnlyList<string> lines, int index) {
        return IsDefinitionLineBlockCandidate(index > 0 ? lines[index - 1] : null)
               || IsDefinitionLineBlockCandidate(index + 1 < lines.Count ? lines[index + 1] : null);
    }

    private static bool HasDefinitionContinuation(IReadOnlyList<string> lines, int index) {
        if (lines == null || index < 0 || index >= lines.Count) {
            return false;
        }

        var line = lines[index] ?? string.Empty;
        int continuationIndent = CountLeadingIndentColumns(line) + 2;
        for (int i = index + 1; i < lines.Count; i++) {
            var next = lines[i] ?? string.Empty;
            if (string.IsNullOrWhiteSpace(next)) {
                continue;
            }

            return CountLeadingIndentColumns(next) >= continuationIndent;
        }

        return false;
    }

    private static bool IsDefinitionLineBlockCandidate(string? line) {
        if (string.IsNullOrWhiteSpace(line)) return false;
        string safeLine = line!;

        int leading = 0;
        while (leading < safeLine.Length && safeLine[leading] == ' ') leading++;
        if (leading >= 4) return false;
        if (leading < safeLine.Length && safeLine[leading] == '\t') return false;

        return IsDefinitionLine(safeLine);
    }

    private static bool TryGetDefinitionSeparator(string line, out int idx) {
        idx = -1;
        if (string.IsNullOrWhiteSpace(line)) return false;
        int start = 0;
        while (start < line.Length) {
            int pos = line.IndexOf(':', start);
            if (pos < 0) return false;
            if (pos > 0 && pos + 1 < line.Length && line[pos + 1] == ' ') {
                var term = line.Substring(0, pos).Trim();
                if (LooksLikeDefinitionTerm(term)) {
                    idx = pos;
                    return true;
                }
            }
            start = pos + 1;
        }
        return false;
    }

    private static bool LooksLikeDefinitionTerm(string term) {
        if (string.IsNullOrWhiteSpace(term)) return false;
        return !ContainsLiteralAutolinkLikeToken(term);
    }

    private static bool ContainsLiteralAutolinkLikeToken(string text) {
        foreach (var rawToken in text.Split([' ', '\t'], StringSplitOptions.RemoveEmptyEntries)) {
            if (LooksLikeMarkdownLinkToken(rawToken)) continue;

            var token = rawToken
                .TrimStart('(', '[', '{', '"', '\'')
                .TrimEnd(')', ']', '}', '"', '\'', '.', ',', ';', '!', '?');
            if (string.IsNullOrWhiteSpace(token)) continue;

            if (token[0] == '<' &&
                TryParseAngleAutolink(token, 0, out int angleConsumed, out _, out _) &&
                angleConsumed == token.Length) {
                return true;
            }

            if ((token[0] == 'h' || token[0] == 'H') &&
                StartsWithHttp(token, 0, out int httpEnd) &&
                httpEnd == token.Length) {
                return true;
            }

            if ((token[0] == 'w' || token[0] == 'W') &&
                StartsWithWww(token, 0, out int wwwEnd) &&
                wwwEnd == token.Length) {
                return true;
            }

            if (IsEmailStartChar(token[0]) &&
                TryConsumePlainEmail(token, 0, out int emailEnd, out _) &&
                emailEnd == token.Length) {
                return true;
            }
        }

        return false;
    }

    private static bool LooksLikeMarkdownLinkToken(string token) {
        if (string.IsNullOrWhiteSpace(token)) return false;

        int start = token[0] == '!' ? 1 : 0;
        if (start >= token.Length || token[start] != '[') return false;

        int closeLabel = token.IndexOf(']', start + 1);
        if (closeLabel < 0 || closeLabel + 1 >= token.Length) return false;

        return (token[closeLabel + 1] == '(' && token[token.Length - 1] == ')') ||
               (token[closeLabel + 1] == '[' && token[token.Length - 1] == ']');
    }

    private static bool IsOrderedListLine(string line, out int number, out string content) {
        number = 0;
        content = string.Empty;
        if (!TryGetOrderedListMarkerInfo(line, out _, out number, out int contentStartIndex)) return false;
        content = line.Substring(contentStartIndex);
        return true;
    }

    private static bool IsOrderedListLine(string line, out int level, out int number, out string content) {
        level = 0;
        number = 0;
        content = string.Empty;
        if (!TryGetOrderedListMarkerInfo(line, out int spaces, out number, out int contentStartIndex)) return false;
        content = line.Substring(contentStartIndex);
        level = spaces / 2;
        return true;
    }

    private static bool IsUnorderedListLine(string line, out bool isTask, out bool done, out string content) {
        isTask = false;
        done = false;
        content = string.Empty;
        if (!TryGetUnorderedListMarkerInfo(line, out _, out int contentStartIndex)) return false;

        var c = line.Substring(contentStartIndex);
        if (c.StartsWith("[ ]", StringComparison.Ordinal)) {
            isTask = true;
            done = false;
            content = c.Length > 3 && c[2] == ']' && c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c;
            return true;
        }

        if (c.StartsWith("[x]", StringComparison.OrdinalIgnoreCase)) {
            isTask = true;
            done = true;
            content = c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c;
            return true;
        }

        content = c;
        return true;
    }

    private static bool IsUnorderedListLine(string line, out int level, out bool isTask, out bool done, out string content) {
        level = 0;
        isTask = false;
        done = false;
        content = string.Empty;
        if (!TryGetUnorderedListMarkerInfo(line, out int spaces, out int contentStartIndex)) return false;

        string c = line.Substring(contentStartIndex);
        if (c.StartsWith("[ ]", StringComparison.Ordinal)) {
            isTask = true;
            done = false;
            content = c.Length > 3 && c[2] == ']' && c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c;
            level = spaces / 2;
            return true;
        }

        if (c.StartsWith("[x]", StringComparison.OrdinalIgnoreCase)) {
            isTask = true;
            done = true;
            content = c.Length > 4 && c[3] == ' ' ? c.Substring(4) : c;
            level = spaces / 2;
            return true;
        }

        content = c;
        level = spaces / 2;
        return true;
    }

    private static string GetUnorderedListItemContent(string line) {
        return TryGetUnorderedListMarkerInfo(line, out _, out int contentStartIndex)
            ? line.Substring(contentStartIndex)
            : string.Empty;
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

    private static int GetListContinuationIndent(string line) {
        if (string.IsNullOrEmpty(line)) return 0;
        if (TryGetOrderedListMarkerInfo(line, out int orderedLeadingSpaces, out _, out int orderedContentStartIndex)) {
            if (string.IsNullOrWhiteSpace(line.Substring(orderedContentStartIndex))
                && TryGetOrderedListMarkerWidth(line, orderedLeadingSpaces, out int orderedMarkerWidth)) {
                return orderedLeadingSpaces + orderedMarkerWidth + 1;
            }

            return orderedContentStartIndex;
        }

        if (TryGetUnorderedListMarkerInfo(line, out int unorderedLeadingSpaces, out int unorderedContentStartIndex)) {
            if (string.IsNullOrWhiteSpace(line.Substring(unorderedContentStartIndex))) {
                return unorderedLeadingSpaces + 2;
            }

            return unorderedContentStartIndex;
        }

        int spaces = CountLeadingSpaces(line);
        return spaces + 2;
    }

    private static int GetRelativeListItemLevel(List<int>? continuationIndentsByLevel, string line) {
        if (continuationIndentsByLevel == null || continuationIndentsByLevel.Count == 0 || string.IsNullOrEmpty(line)) {
            return 0;
        }

        int indentColumns = CountLeadingIndentColumns(line);
        for (int level = continuationIndentsByLevel.Count - 1; level >= 0; level--) {
            if (indentColumns >= continuationIndentsByLevel[level]) {
                return level + 1;
            }
        }

        return 0;
    }

    private static void TrackListItemContinuationIndent(List<int> continuationIndentsByLevel, int level, int continuationIndent) {
        if (continuationIndentsByLevel == null) {
            return;
        }

        while (continuationIndentsByLevel.Count > level) {
            continuationIndentsByLevel.RemoveAt(continuationIndentsByLevel.Count - 1);
        }

        if (continuationIndentsByLevel.Count == level) {
            continuationIndentsByLevel.Add(continuationIndent);
            return;
        }

        continuationIndentsByLevel[level] = continuationIndent;
    }

    private static int GetTaskMarkerConsumedColumns(string content) {
        if (string.IsNullOrEmpty(content)) return 0;
        if (content.StartsWith("[ ]", StringComparison.Ordinal)) {
            return content.Length > 4 && content[3] == ' ' ? 4 : 0;
        }

        if (content.StartsWith("[x]", StringComparison.OrdinalIgnoreCase)) {
            return content.Length > 4 && content[3] == ' ' ? 4 : 0;
        }

        return 0;
    }

    private static bool TryGetRawListItemContentAfterMarker(string line, out string content) {
        content = string.Empty;
        if (string.IsNullOrEmpty(line)) return false;
        if (TryGetOrderedListMarkerInfo(line, out _, out _, out int orderedContentStartIndex)) {
            content = line.Substring(orderedContentStartIndex);
            return true;
        }

        if (TryGetUnorderedListMarkerInfo(line, out _, out int unorderedContentStartIndex)) {
            content = line.Substring(unorderedContentStartIndex);
            return true;
        }

        return false;
    }

    private static bool TryGetOrderedListMarkerInfo(string line, out int leadingSpaces, out int number, out int contentStartIndex) {
        return TryGetOrderedListMarkerInfo(line, out leadingSpaces, out number, out contentStartIndex, out _);
    }

    private static bool TryGetOrderedListMarkerInfo(string line, out int leadingSpaces, out int number, out int contentStartIndex, out char delimiter) {
        leadingSpaces = 0;
        number = 0;
        contentStartIndex = 0;
        delimiter = '\0';
        if (string.IsNullOrEmpty(line)) return false;

        while (leadingSpaces < line.Length && line[leadingSpaces] == ' ') leadingSpaces++;

        int digitsStart = leadingSpaces;
        int digitsEnd = digitsStart;
        while (digitsEnd < line.Length && char.IsDigit(line[digitsEnd])) digitsEnd++;
        if (digitsEnd == digitsStart) return false;
        if (digitsEnd - digitsStart > 9) return false;
        if (digitsEnd >= line.Length || (line[digitsEnd] != '.' && line[digitsEnd] != ')')) return false;
        delimiter = line[digitsEnd];
        if (!TryGetListContentStartIndex(line, digitsEnd, out contentStartIndex)) return false;
        if (!int.TryParse(line.Substring(digitsStart, digitsEnd - digitsStart), NumberStyles.Integer, CultureInfo.InvariantCulture, out number)) number = 1;
        return true;
    }

    private static bool TryGetUnorderedListMarkerInfo(string line, out int leadingSpaces, out int contentStartIndex) {
        return TryGetUnorderedListMarkerInfo(line, out leadingSpaces, out contentStartIndex, out _);
    }

    private static bool TryGetUnorderedListMarkerInfo(string line, out int leadingSpaces, out int contentStartIndex, out char marker) {
        leadingSpaces = 0;
        contentStartIndex = 0;
        marker = '\0';
        if (string.IsNullOrEmpty(line)) return false;

        while (leadingSpaces < line.Length && line[leadingSpaces] == ' ') leadingSpaces++;
        if (leadingSpaces >= line.Length) return false;

        marker = line[leadingSpaces];
        if (marker != '-' && marker != '*' && marker != '+') return false;
        return TryGetListContentStartIndex(line, leadingSpaces, out contentStartIndex);
    }

    private static bool TryGetListContentStartIndex(string line, int markerIndex, out int contentStartIndex) {
        contentStartIndex = 0;
        int paddingStart = markerIndex + 1;
        if (paddingStart >= line.Length) {
            contentStartIndex = line.Length;
            return true;
        }

        int paddingColumns = 0;
        int cursor = paddingStart;
        while (cursor < line.Length) {
            char ch = line[cursor];
            if (ch == ' ' && paddingColumns < 4) {
                paddingColumns++;
                cursor++;
                continue;
            }

            if (ch == '\t' && paddingColumns == 0) {
                contentStartIndex = cursor + 1;
                return true;
            }

            break;
        }

        if (cursor >= line.Length) {
            contentStartIndex = line.Length;
            return true;
        }

        if (paddingColumns == 0) return false;
        contentStartIndex = cursor;
        return true;
    }

    private static bool TryGetIndentedCodeListLead(string line, out int continuationIndent, out string content, out int startColumn) {
        continuationIndent = 0;
        content = string.Empty;
        startColumn = 1;
        if (string.IsNullOrEmpty(line)) return false;

        int leadingSpaces = 0;
        while (leadingSpaces < line.Length && line[leadingSpaces] == ' ') leadingSpaces++;
        if (leadingSpaces >= line.Length) return false;

        int markerWidth;
        if (TryGetOrderedListMarkerWidth(line, leadingSpaces, out markerWidth)) {
            if (!HasIndentedCodePaddingAfterMarker(line, leadingSpaces + markerWidth - 1)) return false;
        } else {
            char marker = line[leadingSpaces];
            if (marker != '-' && marker != '*' && marker != '+') return false;
            markerWidth = 1;
            if (!HasIndentedCodePaddingAfterMarker(line, leadingSpaces)) return false;
        }

        continuationIndent = leadingSpaces + markerWidth + 1;
        if (continuationIndent >= line.Length) return false;

        content = line.Substring(continuationIndent);
        startColumn = continuationIndent + 1;
        return CountLeadingIndentColumns(content) >= 4;
    }

    private static bool TryGetOrderedListMarkerWidth(string line, int leadingSpaces, out int markerWidth) {
        markerWidth = 0;
        if (string.IsNullOrEmpty(line) || leadingSpaces >= line.Length) return false;

        int digitsEnd = leadingSpaces;
        while (digitsEnd < line.Length && char.IsDigit(line[digitsEnd])) digitsEnd++;
        if (digitsEnd == leadingSpaces || digitsEnd - leadingSpaces > 9) return false;
        if (digitsEnd >= line.Length || (line[digitsEnd] != '.' && line[digitsEnd] != ')')) return false;

        markerWidth = digitsEnd - leadingSpaces + 1;
        return true;
    }

    private static bool HasIndentedCodePaddingAfterMarker(string line, int markerEndIndex) {
        int paddingStart = markerEndIndex + 1;
        if (paddingStart >= line.Length || line[paddingStart] != ' ') return false;

        int spaces = 0;
        int cursor = paddingStart;
        while (cursor < line.Length && line[cursor] == ' ') {
            spaces++;
            cursor++;
        }

        return spaces >= 5;
    }

    private static int GetListLeadContentStartColumn(string line, bool stripTaskMarker = false) {
        int startColumn = GetListContinuationIndent(line) + 1;
        if (!stripTaskMarker) return startColumn;

        return TryGetRawListItemContentAfterMarker(line, out string content)
            ? startColumn + GetTaskMarkerConsumedColumns(content)
            : startColumn;
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
