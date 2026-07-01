namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static bool IsReferenceDefinitionParagraphContinuationLine(string[] lines, int index, MarkdownReaderOptions options) {
        if (lines == null || index < 0 || index >= lines.Length) {
            return false;
        }

        var line = lines[index];
        if (string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        if (IsAtxHeading(line, out _, out _) ||
            IsCodeFenceOpen(line, out _, out _, out _) ||
            StartsTable(lines, index, options) ||
            IsParagraphInterruptingThematicBreakLine(line) ||
            IsParagraphInterruptingUnorderedListLine(line) ||
            IsParagraphInterruptingOrderedListLine(line, options) ||
            (options.Callouts && IsCalloutHeader(line, options, out _, out _)) ||
            IsQuoteStarter(line) ||
            HtmlBlockParser.IsParagraphInterruptingHtmlBlockStart(line, options) ||
            TryParseReferenceLinkDefinition(lines, index, options, out _, out _, out _, out _) ||
            IsFootnoteDefinitionStarterForParagraphInterruption(line, options) ||
            (options.StandaloneImageBlocks && IsImageLine(line))) {
            return false;
        }

        return true;
    }

    private static bool IsFootnoteDefinitionStarterForParagraphInterruption(string line, MarkdownReaderOptions options) {
        if (options?.Footnotes != true || string.IsNullOrWhiteSpace(line)) {
            return false;
        }

        int leading = 0;
        while (leading < line.Length && line[leading] == ' ') {
            leading++;
        }

        if (leading >= 4 || (leading < line.Length && line[leading] == '\t')) {
            return false;
        }

        var trimmed = line.TrimStart();
        if (!(trimmed.Length > 4 && trimmed[0] == '[' && trimmed[1] == '^')) {
            return false;
        }

        int rb = trimmed.IndexOf(']');
        return rb >= 2
               && rb + 1 < trimmed.Length
               && trimmed[rb + 1] == ':';
    }
}
