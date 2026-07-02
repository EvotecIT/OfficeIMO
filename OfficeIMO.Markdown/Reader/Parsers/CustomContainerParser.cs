namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static bool IsCustomContainerOpeningLine(string? line, MarkdownReaderOptions options) =>
        options.CustomContainers && CustomContainerParser.IsOpeningLine(line);

    internal sealed class CustomContainerParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.CustomContainers || lines == null || i < 0 || i >= lines.Length) {
                return false;
            }

            if (!TryParseOpeningFence(lines[i], out var opening)) {
                return false;
            }

            var bodyStart = i + 1;
            var closingIndex = -1;
            var closingFenceLength = opening.FenceLength;
            var nestedFenceLengths = new Stack<int>();
            for (var candidate = bodyStart; candidate < lines.Length; candidate++) {
                if (nestedFenceLengths.Count > 0 &&
                    TryParseClosingFence(lines[candidate], nestedFenceLengths.Peek(), out _)) {
                    nestedFenceLengths.Pop();
                    continue;
                }

                if (TryParseClosingFence(lines[candidate], opening.FenceLength, out closingFenceLength)) {
                    if (nestedFenceLengths.Count == 0) {
                        closingIndex = candidate;
                        break;
                    }

                    continue;
                }

                if (TryParseOpeningFence(lines[candidate], out var nestedOpening) &&
                    nestedOpening.FenceLength >= opening.FenceLength) {
                    nestedFenceLengths.Push(nestedOpening.FenceLength);
                }
            }

            var bodyLineCount = (closingIndex < 0 ? lines.Length : closingIndex) - bodyStart;
            IReadOnlyList<IMarkdownBlock> children;
            IReadOnlyList<MarkdownSyntaxNode> syntaxChildren;
            if (bodyLineCount > 0) {
                var sourceLines = new List<MarkdownSourceLineSlice>(bodyLineCount);
                for (var lineIndex = bodyStart; lineIndex < bodyStart + bodyLineCount; lineIndex++) {
                    sourceLines.Add(new MarkdownSourceLineSlice(
                        lines[lineIndex] ?? string.Empty,
                        state.SourceLineOffset + lineIndex + 1,
                        1));
                }

                (children, syntaxChildren) = ParseNestedMarkdownBlocks(sourceLines, options, state);
            } else {
                children = Array.Empty<IMarkdownBlock>();
                syntaxChildren = Array.Empty<MarkdownSyntaxNode>();
            }

            var block = new CustomContainerBlock(opening.Info, children, opening.FenceLength) {
                ClosingFenceLength = closingFenceLength,
                SyntaxChildren = syntaxChildren.Count > 0 ? syntaxChildren : null,
                OpeningFenceSourceSpan = CreateSpan(
                    state,
                    state.SourceLineOffset + i + 1,
                    opening.FenceStartColumn,
                    state.SourceLineOffset + i + 1,
                    opening.FenceStartColumn + opening.FenceLength - 1)
            };

            if (opening.InfoStartColumn > 0 && opening.Info.Length > 0) {
                var openingLine = lines[i] ?? string.Empty;
                var infoStartIndex = opening.InfoStartColumn - 1;
                var infoStartColumn = AdvanceSourceColumn(1, openingLine, infoStartIndex);
                var infoEndColumn = AdvanceSourceColumn(1, openingLine, infoStartIndex + opening.Info.Length) - 1;
                block.InfoSourceSpan = CreateSpan(
                    state,
                    state.SourceLineOffset + i + 1,
                    infoStartColumn,
                    state.SourceLineOffset + i + 1,
                    infoEndColumn);
            }

            if (closingIndex >= 0) {
                var closingLine = lines[closingIndex] ?? string.Empty;
                var closingIndent = CountLeadingSpaces(closingLine);
                block.ClosingFenceSourceSpan = CreateSpan(
                    state,
                    state.SourceLineOffset + closingIndex + 1,
                    closingIndent + 1,
                    state.SourceLineOffset + closingIndex + 1,
                    closingIndent + closingFenceLength);
            }

            doc.Add(block);
            i = closingIndex < 0 ? lines.Length : closingIndex + 1;
            return true;
        }

        internal static bool IsOpeningLine(string? line) =>
            TryParseOpeningFence(line, out _);

        internal static bool TryGetContainerLineCount(IReadOnlyList<string> lines, int startIndex, out int lineCount) {
            lineCount = 0;
            if (lines == null || startIndex < 0 || startIndex >= lines.Count) {
                return false;
            }

            if (!TryParseOpeningFence(lines[startIndex], out var opening)) {
                return false;
            }

            var bodyStart = startIndex + 1;
            var nestedFenceLengths = new Stack<int>();
            for (var candidate = bodyStart; candidate < lines.Count; candidate++) {
                if (nestedFenceLengths.Count > 0 &&
                    TryParseClosingFence(lines[candidate], nestedFenceLengths.Peek(), out _)) {
                    nestedFenceLengths.Pop();
                    continue;
                }

                if (TryParseClosingFence(lines[candidate], opening.FenceLength, out _)) {
                    if (nestedFenceLengths.Count == 0) {
                        lineCount = candidate - startIndex + 1;
                        return true;
                    }

                    continue;
                }

                if (TryParseOpeningFence(lines[candidate], out var nestedOpening) &&
                    nestedOpening.FenceLength >= opening.FenceLength) {
                    nestedFenceLengths.Push(nestedOpening.FenceLength);
                }
            }

            lineCount = lines.Count - startIndex;
            return true;
        }

        private static bool TryParseOpeningFence(string? line, out CustomContainerFence opening) {
            opening = default;
            line ??= string.Empty;
            if (CountLeadingIndentColumns(line) > 3) {
                return false;
            }

            var indent = CountLeadingSpaces(line);
            var index = indent;
            var fenceLength = CountColonRun(line, index);
            if (fenceLength < 3) {
                return false;
            }

            var infoStart = index + fenceLength;
            while (infoStart < line.Length && char.IsWhiteSpace(line[infoStart])) {
                infoStart++;
            }

            var infoEnd = line.Length;
            while (infoEnd > infoStart && char.IsWhiteSpace(line[infoEnd - 1])) {
                infoEnd--;
            }

            var info = infoEnd > infoStart
                ? line.Substring(infoStart, infoEnd - infoStart)
                : string.Empty;

            opening = new CustomContainerFence(
                fenceLength,
                indent + 1,
                info,
                info.Length == 0 ? 0 : infoStart + 1);
            return true;
        }

        private static bool TryParseClosingFence(string? line, int openingFenceLength, out int closingFenceLength) {
            closingFenceLength = 0;
            line ??= string.Empty;
            if (CountLeadingIndentColumns(line) > 3) {
                return false;
            }

            var indent = CountLeadingSpaces(line);
            var fenceLength = CountColonRun(line, indent);
            if (fenceLength < openingFenceLength) {
                return false;
            }

            for (var index = indent + fenceLength; index < line.Length; index++) {
                if (!char.IsWhiteSpace(line[index])) {
                    return false;
                }
            }

            closingFenceLength = fenceLength;
            return true;
        }

        private static int CountColonRun(string line, int index) {
            var count = 0;
            while (index + count < line.Length && line[index + count] == ':') {
                count++;
            }

            return count;
        }

        private static int AdvanceSourceColumn(int startColumn, string? text, int endExclusive) {
            var column = Math.Max(1, startColumn);
            var boundedEnd = Math.Max(0, Math.Min(endExclusive, text?.Length ?? 0));
            for (var i = 0; i < boundedEnd; i++) {
                column = MarkdownSourceColumns.AdvanceColumn(column, text![i]);
            }

            return column;
        }

        private readonly struct CustomContainerFence {
            public CustomContainerFence(int fenceLength, int fenceStartColumn, string info, int infoStartColumn) {
                FenceLength = fenceLength;
                FenceStartColumn = fenceStartColumn;
                Info = info;
                InfoStartColumn = infoStartColumn;
            }

            public int FenceLength { get; }
            public int FenceStartColumn { get; }
            public string Info { get; }
            public int InfoStartColumn { get; }
        }
    }
}
