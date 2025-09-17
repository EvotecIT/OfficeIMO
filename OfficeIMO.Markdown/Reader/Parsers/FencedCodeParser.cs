namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    internal sealed class FencedCodeParser : IMarkdownBlockParser {
        public bool TryParse(string[] lines, ref int i, MarkdownReaderOptions options, MarkdownDoc doc, MarkdownReaderState state) {
            if (!options.FencedCode) return false;
            if (!IsCodeFenceOpen(lines[i], out string language)) return false;
            int j = i + 1;
            var code = new System.Text.StringBuilder();
            while (j < lines.Length && !IsCodeFenceClose(lines[j])) { code.AppendLine(lines[j]); j++; }
            if (j < lines.Length && IsCodeFenceClose(lines[j])) j++;
            var block = new CodeBlock(language, code.ToString().TrimEnd('\n'));
            if (j < lines.Length && TryParseCaption(lines[j], out var cap)) { block.Caption = cap; j++; }
            doc.Add(block);
            i = j; return true;
        }
    }
}
