namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    internal sealed class FencedCodeParser : MarkdownSourceLinesBlockParser {
        public override bool TryParse(MarkdownSourceLines lines, ref int i, MarkdownLoadOptions options, MarkdownDocument doc, MarkdownReaderState state) {
            if (!options.FencedCode) return false;
            if (!IsCodeFenceOpen(lines, i, out string language, out char fenceChar, out int fenceLength, out int fenceIndentColumns, out int infoPaddingColumns, out int infoPaddingCharacters)) return false;
            int j = i + 1;
            int lineLimit = lines.Length - (state.HasSyntheticTerminalLine ? 1 : 0);
            var code = MarkdownStringBuilderCache.Acquire();
            while (j < lineLimit && !IsCodeFenceClose(lines, j, fenceChar, fenceLength)) {
                if (fenceIndentColumns == 0) {
                    lines.GetSlice(j, out var lineText, out int lineStart, out int lineEnd);
                    code.Append(lineText, lineStart, lineEnd - lineStart).AppendLine();
                } else {
                    code.AppendLine(StripLeadingIndentColumns(lines[j], fenceIndentColumns));
                }
                j++;
            }
            var contentLineCount = j - i - 1;
            bool hasClosingFence = false;
            int closingFenceIndentColumns = 0;
            int closingFenceLength = fenceLength;
            if (j < lineLimit && TryGetCodeFenceCloseInfo(lines, j, fenceChar, fenceLength, out closingFenceIndentColumns, out closingFenceLength)) {
                hasClosingFence = true;
                j++;
            }
            string? caption = null;
            if (j < lineLimit && lines.GetFirstNonWhiteSpaceCharacter(j) == '_' && TryParseCaption(lines[j], out var cap)) { caption = cap; j++; }
            var block = CreateParsedFencedBlock(
                language,
                ToStringWithoutSingleTrailingLineEnding(code),
                isFenced: true,
                caption,
                options,
                fenceIndentColumns,
                fenceLength,
                infoPaddingColumns,
                infoPaddingCharacters,
                fenceChar,
                hasClosingFence,
                closingFenceIndentColumns,
                closingFenceLength,
                contentLineCount);
            doc.Add(block);
            i = j; return true;
        }
    }
}
