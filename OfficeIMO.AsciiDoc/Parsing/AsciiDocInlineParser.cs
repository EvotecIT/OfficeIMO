namespace OfficeIMO.AsciiDoc;

internal sealed class AsciiDocInlineParser {
    private readonly AsciiDocSyntaxFactory _factory;
    private readonly AsciiDocParseOptions _options;
    private int _nodeCount;

    internal AsciiDocInlineParser(AsciiDocSyntaxFactory factory, AsciiDocParseOptions options) {
        _factory = factory;
        _options = options;
    }

    internal AsciiDocInlineSequence Parse(int start, int end) => ParseRange(start, end, 0);

    private AsciiDocInlineSequence ParseRange(int start, int end, int depth) {
        if (depth > _options.MaximumInlineNestingDepth) throw new InvalidDataException("AsciiDoc source exceeds MaximumInlineNestingDepth.");
        var items = new List<AsciiDocInline>();
        var syntax = new List<AsciiDocSyntaxNode>();
        int textStart = start;
        int index = start;
        while (index < end) {
            // An escape consumes itself and the following character. Because the parser
            // advances past both characters, it never lands on an escaped backslash and
            // does not need to rescan the preceding run for every slash.
            if (_factory.Source.Text[index] == '\\') { index += Math.Min(2, end - index); continue; }
            AsciiDocInline? item = TryParseExplicit(index, end, depth, out int next);
            if (item == null) { index++; continue; }
            AddText(textStart, index, items, syntax);
            Add(item, items, syntax);
            index = next;
            textStart = index;
        }
        AddText(textStart, end, items, syntax);
        AsciiDocSyntaxNode sequenceSyntax = _factory.Node(AsciiDocSyntaxKind.InlineSequence, start, end, syntax);
        return new AsciiDocInlineSequence(sequenceSyntax, items);
    }

    private AsciiDocInline? TryParseExplicit(int start, int end, int depth, out int next) {
        next = start;
        string source = _factory.Source.Text;
        if (source[start] == '<' && start + 1 < end && source[start + 1] == '<') return ParseReference(start, end, false, out next);
        if (source[start] == '[' && start + 1 < end && source[start + 1] == '[') return ParseReference(start, end, true, out next);
        if (source[start] == '{') return ParseAttributeReference(start, end, out next);
        if (AsciiDocText.IsAsciiLetter(source[start])) {
            AsciiDocInline? macro = ParseMacro(start, end, out next);
            if (macro != null) return macro;
        }
        if (source[start] == '+') return ParsePassthrough(start, end, out next);
        return ParseFormatting(start, end, depth, out next);
    }

    private AsciiDocInline? ParseReference(int start, int end, bool anchor, out int next) {
        string source = _factory.Source.Text;
        string close = anchor ? "]]" : ">>";
        int closing = FindToken(source, close, start + 2, end);
        if (closing < 0) { next = start; return null; }
        int contentStart = start + 2;
        string content = source.Substring(contentStart, closing - contentStart);
        int comma = FindUnescaped(content, ',');
        string target = comma < 0 ? content : content.Substring(0, comma);
        string? label = comma < 0 ? null : content.Substring(comma + 1);
        next = closing + 2;
        AsciiDocSyntaxKind kind = anchor ? AsciiDocSyntaxKind.InlineAnchor : AsciiDocSyntaxKind.InlineCrossReference;
        AsciiDocSyntaxNode syntax = _factory.Node(kind, start, next);
        return anchor
            ? new AsciiDocAnchorInline(syntax, target, label)
            : new AsciiDocCrossReferenceInline(syntax, target, label);
    }

    private AsciiDocInline? ParseAttributeReference(int start, int end, out int next) {
        string source = _factory.Source.Text;
        int closing = FindUnescaped(source, '}', start + 1, end);
        if (closing < 0) { next = start; return null; }
        string name = source.Substring(start + 1, closing - start - 1);
        if (!AsciiDocText.IsAttributeName(name)) { next = start; return null; }
        next = closing + 1;
        return new AsciiDocAttributeReferenceInline(_factory.Node(AsciiDocSyntaxKind.InlineAttributeReference, start, next), name);
    }

    private AsciiDocInline? ParseMacro(int start, int end, out int next) {
        string source = _factory.Source.Text;
        int nameEnd = start + 1;
        while (nameEnd < end && (AsciiDocText.IsAsciiLetter(source[nameEnd]) || char.IsDigit(source[nameEnd]) || source[nameEnd] == '_' || source[nameEnd] == '-')) nameEnd++;
        if (nameEnd >= end || source[nameEnd] != ':') { next = start; return null; }
        int open = FindUnescaped(source, '[', nameEnd + 1, end);
        if (open < 0) { next = start; return null; }
        int close = FindMatchingBracket(source, open, end);
        if (close < 0) { next = start; return null; }
        string name = source.Substring(start, nameEnd - start);
        string target = source.Substring(nameEnd + 1, open - nameEnd - 1);
        string attributes = source.Substring(open + 1, close - open - 1);
        next = close + 1;
        AsciiDocSyntaxKind kind = IsStem(name) ? AsciiDocSyntaxKind.InlineStem : AsciiDocSyntaxKind.InlineMacro;
        AsciiDocSyntaxNode syntax = _factory.Node(kind, start, next);
        return IsStem(name) && target.Length == 0
            ? new AsciiDocStemInline(syntax, name, attributes)
            : new AsciiDocMacroInline(syntax, name, target, attributes);
    }

    private AsciiDocInline? ParsePassthrough(int start, int end, out int next) {
        string source = _factory.Source.Text;
        int length = 1;
        while (length < 3 && start + length < end && source[start + length] == '+') length++;
        string marker = new string('+', length);
        int close = FindToken(source, marker, start + length, end);
        if (close < 0 || close == start + length) { next = start; return null; }
        next = close + length;
        string content = source.Substring(start + length, close - start - length);
        return new AsciiDocPassthroughInline(_factory.Node(AsciiDocSyntaxKind.InlinePassthrough, start, next), marker, content);
    }

    private AsciiDocInline? ParseFormatting(int start, int end, int depth, out int next) {
        string source = _factory.Source.Text;
        if (!TryGetStyle(source[start], out AsciiDocInlineStyle style)) { next = start; return null; }
        int length = start + 1 < end && source[start + 1] == source[start] ? 2 : 1;
        if (length == 1 && !IsConstrainedOpening(source, start, end)) { next = start; return null; }
        string marker = new string(source[start], length);
        int search = start + length;
        int close;
        while ((close = FindToken(source, marker, search, end)) >= 0) {
            if (close > start + length && (length == 2 || IsConstrainedClosing(source, close, length, end))) break;
            search = close + length;
        }
        if (close < 0) { next = start; return null; }
        AsciiDocInlineSequence content = ParseRange(start + length, close, depth + 1);
        next = close + length;
        var children = new List<AsciiDocSyntaxNode> {
            _factory.Node(AsciiDocSyntaxKind.InlineFormattingMarker, start, start + length),
            content.Syntax,
            _factory.Node(AsciiDocSyntaxKind.InlineFormattingMarker, close, next)
        };
        AsciiDocSyntaxNode syntax = _factory.Node(AsciiDocSyntaxKind.InlineFormatted, start, next, children);
        return new AsciiDocFormattedInline(syntax, style, marker, content);
    }

    private void AddText(int start, int end, List<AsciiDocInline> items, List<AsciiDocSyntaxNode> syntaxNodes) {
        if (end <= start) return;
        AsciiDocSyntaxNode syntax = _factory.Node(AsciiDocSyntaxKind.InlineText, start, end);
        Add(new AsciiDocTextInline(syntax, syntax.OriginalText), items, syntaxNodes);
    }

    private void Add(AsciiDocInline item, List<AsciiDocInline> items, List<AsciiDocSyntaxNode> syntaxNodes) {
        _nodeCount++;
        if (_nodeCount > _options.MaximumInlineNodeCount) throw new InvalidDataException("AsciiDoc source exceeds MaximumInlineNodeCount.");
        items.Add(item);
        syntaxNodes.Add(item.Syntax);
    }

    private static bool TryGetStyle(char marker, out AsciiDocInlineStyle style) {
        switch (marker) {
            case '*': style = AsciiDocInlineStyle.Strong; return true;
            case '_': style = AsciiDocInlineStyle.Emphasis; return true;
            case '`': style = AsciiDocInlineStyle.Monospace; return true;
            case '#': style = AsciiDocInlineStyle.Highlight; return true;
            case '~': style = AsciiDocInlineStyle.Subscript; return true;
            case '^': style = AsciiDocInlineStyle.Superscript; return true;
            default: style = default; return false;
        }
    }

    private static bool IsConstrainedOpening(string source, int start, int end) {
        int inner = start + 1;
        if (inner >= end || char.IsWhiteSpace(source[inner])) return false;
        if (start == 0) return true;
        char before = source[start - 1];
        return !IsWord(before) && before != ':' && before != ';' && before != '}';
    }

    private static bool IsConstrainedClosing(string source, int close, int length, int end) {
        if (close <= 0 || char.IsWhiteSpace(source[close - 1])) return false;
        int afterIndex = close + length;
        if (afterIndex >= end) return true;
        char after = source[afterIndex];
        return !IsWord(after) && after != ':' && after != ';' && after != '{';
    }

    private static bool IsWord(char value) => char.IsLetterOrDigit(value) || value == '_';

    private static bool IsStem(string name) =>
        string.Equals(name, "stem", StringComparison.Ordinal) ||
        string.Equals(name, "latexmath", StringComparison.Ordinal) ||
        string.Equals(name, "asciimath", StringComparison.Ordinal);

    private static int FindToken(string source, string token, int start, int end) {
        for (int index = start; index + token.Length <= end; index++) {
            if (source[index] == '\\') { index++; continue; }
            bool match = true;
            for (int offset = 0; offset < token.Length; offset++) {
                if (source[index + offset] != token[offset]) { match = false; break; }
            }
            if (match) return index;
        }
        return -1;
    }

    private static int FindUnescaped(string source, char token, int start, int end) {
        for (int index = start; index < end; index++) {
            if (source[index] == '\\') { index++; continue; }
            if (source[index] == token) return index;
        }
        return -1;
    }

    private static int FindUnescaped(string source, char token) => FindUnescaped(source, token, 0, source.Length);

    private static int FindMatchingBracket(string source, int open, int end) {
        char quote = '\0';
        for (int index = open + 1; index < end; index++) {
            char current = source[index];
            if (current == '\\') { index++; continue; }
            if (quote != '\0') {
                if (current == quote) quote = '\0';
                continue;
            }
            if (current == '\'' || current == '"') quote = current;
            else if (current == ']') return index;
        }
        return -1;
    }
}
