namespace OfficeIMO.Markdown;

internal static partial class MarkdownReader {
    internal sealed class FencedCodeParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownLoadOptions options, MarkdownDocument doc, MarkdownReaderState state) {
            if (!options.FencedCode) return false;
            if (!IsCodeFenceOpen(lines[i], out string language, out char fenceChar, out int fenceLength, out int fenceIndentColumns, out int infoPaddingColumns, out int infoPaddingCharacters)) return false;
            int j = i + 1;
            int lineLimit = lines.Length - (state.HasSyntheticTerminalLine ? 1 : 0);
            var code = new System.Text.StringBuilder();
            while (j < lineLimit && !IsCodeFenceClose(lines[j], fenceChar, fenceLength)) {
                code.AppendLine(StripLeadingIndentColumns(lines[j], fenceIndentColumns));
                j++;
            }
            var contentLineCount = j - i - 1;
            bool hasClosingFence = false;
            int closingFenceIndentColumns = 0;
            int closingFenceLength = fenceLength;
            if (j < lineLimit && TryGetCodeFenceCloseInfo(lines[j], fenceChar, fenceLength, out closingFenceIndentColumns, out closingFenceLength)) {
                hasClosingFence = true;
                j++;
            }
            string? caption = null;
            if (j < lineLimit && TryParseCaption(lines[j], out var cap)) { caption = cap; j++; }
            var block = CreateParsedFencedBlock(
                language,
                RemoveSingleTrailingLineEnding(code.ToString()),
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
