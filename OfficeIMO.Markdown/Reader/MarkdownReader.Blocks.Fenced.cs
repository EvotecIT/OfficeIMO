using System.Globalization;
using System.Linq;
using System.Text;

namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private static bool IsCodeFenceOpen(string line, out string infoString, out char fenceChar, out int fenceLength) =>
        IsCodeFenceOpen(line, out infoString, out fenceChar, out fenceLength, out _, out _);

    private static bool IsCodeFenceOpen(
        string line,
        out string infoString,
        out char fenceChar,
        out int fenceLength,
        out int fenceIndentColumns,
        out int infoPaddingColumns) {
        infoString = string.Empty;
        fenceChar = '\0';
        fenceLength = 0;
        fenceIndentColumns = 0;
        infoPaddingColumns = 0;
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
        fenceIndentColumns = indent;
        infoPaddingColumns = CountLeadingIndentColumns(parsedInfoString);
        infoString = parsedInfoString.Trim();
        return true;
    }
    private static bool IsCodeFenceClose(string line, char fenceChar, int fenceLength) =>
        TryGetCodeFenceCloseInfo(line, fenceChar, fenceLength, out _, out _);

    private static bool TryGetCodeFenceCloseInfo(string line, char fenceChar, int fenceLength, out int fenceIndentColumns, out int closingFenceLength) {
        fenceIndentColumns = 0;
        closingFenceLength = 0;
        if (line is null) return false;
        int indent = CountLeadingIndentColumns(line);
        if (indent > 3) return false;

        var candidate = indent > 0 ? StripLeadingIndentColumns(line, indent) : line;
        int run = 0;
        while (run < candidate.Length && candidate[run] == fenceChar) run++;
        if (run < Math.Max(3, fenceLength)) return false;
        for (int i = run; i < candidate.Length; i++) {
            if (!char.IsWhiteSpace(candidate[i])) return false;
        }

        fenceIndentColumns = indent;
        closingFenceLength = run;
        return true;
    }

    private static IMarkdownBlock CreateParsedFencedBlock(
        string infoString,
        string content,
        bool isFenced,
        string? caption,
        MarkdownReaderOptions options,
        int fenceIndentColumns = 0,
        int fenceLength = 3,
        int infoPaddingColumns = 0,
        char fenceChar = '`',
        bool hasClosingFence = true,
        int closingFenceIndentColumns = 0,
        int closingFenceLength = 3) {
        var extendedBlock = TryCreateExtendedFencedBlock(options?.FencedBlockExtensions, infoString, content, isFenced, caption);
        if (extendedBlock != null) {
            ApplyFenceSourceInfo(extendedBlock, fenceIndentColumns, fenceLength, infoPaddingColumns, fenceChar, hasClosingFence, closingFenceIndentColumns, closingFenceLength);
            return extendedBlock;
        }

        var codeBlock = new CodeBlock(infoString, content, isFenced) {
            Caption = caption
        };
        codeBlock.SetFenceSourceInfo(fenceIndentColumns, fenceLength, infoPaddingColumns, fenceChar, hasClosingFence, closingFenceIndentColumns, closingFenceLength);
        return codeBlock;
    }

    private static void ApplyFenceSourceInfo(
        IMarkdownBlock block,
        int fenceIndentColumns,
        int fenceLength,
        int infoPaddingColumns,
        char fenceChar,
        bool hasClosingFence,
        int closingFenceIndentColumns,
        int closingFenceLength) {
        switch (block) {
            case CodeBlock code:
                code.SetFenceSourceInfo(fenceIndentColumns, fenceLength, infoPaddingColumns, fenceChar, hasClosingFence, closingFenceIndentColumns, closingFenceLength);
                break;
            case SemanticFencedBlock semantic:
                semantic.SetFenceSourceInfo(fenceIndentColumns, fenceLength, infoPaddingColumns, fenceChar, hasClosingFence, closingFenceIndentColumns, closingFenceLength);
                break;
        }
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

}
