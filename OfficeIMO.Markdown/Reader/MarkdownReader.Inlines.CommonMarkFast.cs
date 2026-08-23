namespace OfficeIMO.Markdown;

public static partial class MarkdownReader {
    private enum SimpleInlineTokenKind {
        Strong,
        Code,
        Link
    }

    private readonly record struct SimpleInlineToken(
        SimpleInlineTokenKind Kind,
        int Start,
        int End,
        int ContentStart,
        int ContentLength,
        int TargetStart = 0,
        int TargetLength = 0);

    private static bool TryParseSimpleCommonMarkInlines(
        string text,
        MarkdownReaderOptions options,
        MarkdownReaderState? state,
        MarkdownInlineSourceMap? sourceMap,
        out InlineSequence sequence) {
        sequence = null!;
        if (state?.CaptureSyntaxTree != false
            || sourceMap != null
            || options.InlineParserExtensions.Count != 0
            || options.InlineTransformExtensions.Count != 0
            || options.Abbreviations
            || options.AutolinkUrls
            || options.AutolinkWwwUrls
            || options.AutolinkBareSchemeUrls
            || options.AutolinkEmails
            || string.IsNullOrEmpty(text)) {
            return false;
        }

        bool allowLiteralBrackets = !options.Footnotes && state.LinkRefs.Count == 0;
        if (!ValidateSimpleCommonMarkInlineTokens(text, options, allowLiteralBrackets)) {
            return false;
        }

        sequence = new InlineSequence { AutoSpacing = false };
        int position = 0;
        int textStart = 0;
        while (position < text.Length) {
            if (!TryReadSimpleCommonMarkInlineToken(text, position, options, out var token)) {
                position++;
                continue;
            }

            AddSimpleTextRun(sequence, text, textStart, token.Start - textStart);
            string content = text.Substring(token.ContentStart, token.ContentLength);
            switch (token.Kind) {
                case SimpleInlineTokenKind.Strong:
                    var nested = new InlineSequence { AutoSpacing = false };
                    nested.AddRaw(new MarkdownTextRun(content));
                    sequence.AddRaw(new BoldSequenceInline(nested));
                    break;

                case SimpleInlineTokenKind.Code:
                    sequence.AddRaw(new CodeSpanInline(content));
                    break;

                case SimpleInlineTokenKind.Link:
                    string target = text.Substring(token.TargetStart, token.TargetLength);
                    string? resolvedTarget = string.IsNullOrWhiteSpace(target)
                        ? string.Empty
                        : ResolveUrl(target, options);
                    if (resolvedTarget == null) {
                        sequence.AddRaw(new MarkdownTextRun(content));
                    } else {
                        var label = new InlineSequence { AutoSpacing = false };
                        label.AddRaw(new MarkdownTextRun(content));
                        sequence.AddRaw(new LinkInline(label, resolvedTarget, title: null));
                    }
                    break;
            }

            position = token.End;
            textStart = position;
        }

        AddSimpleTextRun(sequence, text, textStart, text.Length - textStart);
        return true;
    }

    private static bool ValidateSimpleCommonMarkInlineTokens(
        string text,
        MarkdownReaderOptions options,
        bool allowLiteralBrackets) {
        bool foundToken = false;
        bool foundLiteralBracket = false;
        int position = 0;
        while (position < text.Length) {
            char value = text[position];
            if (value is '*' or '`' or '[') {
                if (!TryReadSimpleCommonMarkInlineToken(text, position, options, out var token)) {
                    if (value == '['
                        && allowLiteralBrackets
                        && position + 2 < text.Length
                        && text[position + 2] == ']'
                        && text[position + 1] is ' ' or 'x' or 'X') {
                        foundLiteralBracket = true;
                        position += 3;
                        continue;
                    }

                    return false;
                }

                foundToken = true;
                position = token.End;
                continue;
            }

            if (value is '\\' or '&' or '\n' or '<' or '!' or '_' or '~' or '=' or '+' or '^') {
                return false;
            }

            position++;
        }

        return foundToken || foundLiteralBracket;
    }

    private static bool TryReadSimpleCommonMarkInlineToken(
        string text,
        int start,
        MarkdownReaderOptions options,
        out SimpleInlineToken token) {
        token = default;
        if (start < 0 || start >= text.Length) {
            return false;
        }

        if (text[start] == '*') {
            if (start + 3 >= text.Length || text[start + 1] != '*') {
                return false;
            }

            int closing = text.IndexOf("**", start + 2, StringComparison.Ordinal);
            if (closing <= start + 2
                || !IsSimpleInlineLiteral(text, start + 2, closing - start - 2)) {
                return false;
            }

            GetDelimiterFlags(text, start, '*', 2, options.CjkFriendlyEmphasis, out bool canOpen, out _);
            GetDelimiterFlags(text, closing, '*', 2, options.CjkFriendlyEmphasis, out _, out bool canClose);
            if (!canOpen || !canClose) {
                return false;
            }

            token = new SimpleInlineToken(
                SimpleInlineTokenKind.Strong,
                start,
                closing + 2,
                start + 2,
                closing - start - 2);
            return true;
        }

        if (text[start] == '`') {
            if (start + 2 >= text.Length || text[start + 1] == '`') {
                return false;
            }

            int closing = text.IndexOf('`', start + 1);
            if (closing <= start + 1
                || char.IsWhiteSpace(text[start + 1])
                || char.IsWhiteSpace(text[closing - 1])
                || ContainsLineBreak(text, start + 1, closing - start - 1)) {
                return false;
            }

            token = new SimpleInlineToken(
                SimpleInlineTokenKind.Code,
                start,
                closing + 1,
                start + 1,
                closing - start - 1);
            return true;
        }

        if (text[start] != '[') {
            return false;
        }

        int labelEnd = text.IndexOf(']', start + 1);
        if (labelEnd <= start + 1
            || labelEnd + 2 >= text.Length
            || text[labelEnd + 1] != '('
            || !IsSimpleInlineLiteral(text, start + 1, labelEnd - start - 1)) {
            return false;
        }

        int targetStart = labelEnd + 2;
        int targetEnd = text.IndexOf(')', targetStart);
        if (targetEnd < targetStart || !IsSimpleInlineLinkTarget(text, targetStart, targetEnd - targetStart)) {
            return false;
        }

        token = new SimpleInlineToken(
            SimpleInlineTokenKind.Link,
            start,
            targetEnd + 1,
            start + 1,
            labelEnd - start - 1,
            targetStart,
            targetEnd - targetStart);
        return true;
    }

    private static bool ContainsLineBreak(string text, int start, int length) {
        int end = Math.Min(text.Length, start + length);
        for (int i = Math.Max(0, start); i < end; i++) {
            if (text[i] is '\r' or '\n') {
                return true;
            }
        }

        return false;
    }

    private static bool IsSimpleInlineLiteral(string text, int start, int length) {
        if (length <= 0 || start < 0 || start + length > text.Length) {
            return false;
        }

        int end = start + length;
        for (int i = start; i < end; i++) {
            char value = text[i];
            if (value is '\\' or '&' or '\r' or '\n' or '<' or '!' or '[' or '`' or '*' or '_' or '~' or '=' or '+' or '^') {
                return false;
            }
        }

        return true;
    }

    private static bool IsSimpleInlineLinkTarget(string text, int start, int length) {
        if (length < 0 || start < 0 || start + length > text.Length) {
            return false;
        }

        int end = start + length;
        for (int i = start; i < end; i++) {
            char value = text[i];
            if (char.IsWhiteSpace(value)
                || value is '(' or ')' or '\\' or '<' or '>' or '"' or '\'' or '&') {
                return false;
            }
        }

        return true;
    }

    private static void AddSimpleTextRun(InlineSequence sequence, string text, int start, int length) {
        if (length > 0) {
            sequence.AddRaw(new MarkdownTextRun(text.Substring(start, length)));
        }
    }
}
