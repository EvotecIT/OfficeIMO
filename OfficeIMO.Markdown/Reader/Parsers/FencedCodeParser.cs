namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class FencedCodeParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.FencedCode) return false;
            if (!IsCodeFenceOpen(lines[i], out string language, out char fenceChar, out int fenceLength)) return false;
            int j = i + 1;
            var code = new System.Text.StringBuilder();
            while (j < lines.Length && !IsCodeFenceClose(lines[j], fenceChar, fenceLength)) { code.AppendLine(lines[j]); j++; }
            if (j < lines.Length && IsCodeFenceClose(lines[j], fenceChar, fenceLength)) j++;
            string? caption = null;
            if (j < lines.Length && TryParseCaption(lines[j], out var cap)) { caption = cap; j++; }
            var block = CreateParsedFencedBlock(language, RemoveSingleTrailingLineEnding(code.ToString()), isFenced: true, caption, options);
            doc.Add(block);
            i = j; return true;
        }
    }
}
