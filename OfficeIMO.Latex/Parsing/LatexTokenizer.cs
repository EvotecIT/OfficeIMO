namespace OfficeIMO.Latex;

/// <summary>Dependency-free TeX-aware tokenizer that never expands commands.</summary>
public static class LatexTokenizer {
    /// <summary>Tokenizes every input character exactly once.</summary>
    public static IReadOnlyList<LatexToken> Tokenize(string source, LatexParseOptions? options = null) {
        if (source == null) throw new ArgumentNullException(nameof(source));
        options ??= new LatexParseOptions();
        Validate(source, options);
        var sourceText = new LatexSourceText(source);
        var tokens = new List<LatexToken>();
        int index = 0;
        while (index < source.Length) {
            if (tokens.Count >= options.MaximumTokenCount) throw new InvalidDataException("LaTeX source exceeds MaximumTokenCount.");
            int start = index;
            char current = source[index];
            LatexTokenKind kind;
            string? value = null;
            if (current == '\\') {
                index++;
                if (index < source.Length && IsControlWordCharacter(source[index])) {
                    int nameStart = index;
                    while (index < source.Length && IsControlWordCharacter(source[index])) index++;
                    value = source.Substring(nameStart, index - nameStart);
                } else if (index < source.Length) {
                    value = source[index].ToString();
                    index++;
                } else {
                    value = string.Empty;
                }
                kind = LatexTokenKind.Command;
            } else if (current == '%') {
                index++;
                while (index < source.Length && source[index] != '\r' && source[index] != '\n') index++;
                kind = LatexTokenKind.Comment;
            } else if (current == '\r' || current == '\n') {
                if (current == '\r' && index + 1 < source.Length && source[index + 1] == '\n') index += 2;
                else index++;
                kind = LatexTokenKind.LineEnding;
            } else if (current == ' ' || current == '\t') {
                index++;
                while (index < source.Length && (source[index] == ' ' || source[index] == '\t')) index++;
                kind = LatexTokenKind.Whitespace;
            } else if (current == '$') {
                index++;
                if (index < source.Length && source[index] == '$') index++;
                kind = LatexTokenKind.MathShift;
            } else if (TryGetSingleKind(current, out kind)) {
                index++;
            } else {
                index++;
                while (index < source.Length && !IsSpecial(source[index])) index++;
                kind = LatexTokenKind.Text;
            }
            string text = source.Substring(start, index - start);
            tokens.Add(new LatexToken(kind, text, value, sourceText.CreateSpan(start, index)));
        }
        return tokens;
    }

    private static bool TryGetSingleKind(char value, out LatexTokenKind kind) {
        switch (value) {
            case '{': kind = LatexTokenKind.OpenBrace; return true;
            case '}': kind = LatexTokenKind.CloseBrace; return true;
            case '[': kind = LatexTokenKind.OpenBracket; return true;
            case ']': kind = LatexTokenKind.CloseBracket; return true;
            case '&': kind = LatexTokenKind.AlignmentTab; return true;
            case '^': kind = LatexTokenKind.Superscript; return true;
            case '_': kind = LatexTokenKind.Subscript; return true;
            case '#': kind = LatexTokenKind.Parameter; return true;
            case '~': kind = LatexTokenKind.NonBreakingSpace; return true;
            default: kind = default; return false;
        }
    }

    private static bool IsSpecial(char value) =>
        value == '\\' || value == '%' || value == '\r' || value == '\n' || value == ' ' || value == '\t' ||
        value == '$' || value == '{' || value == '}' || value == '[' || value == ']' || value == '&' ||
        value == '^' || value == '_' || value == '#' || value == '~';

    private static bool IsControlWordCharacter(char value) =>
        (value >= 'a' && value <= 'z') || (value >= 'A' && value <= 'Z') || value == '@';

    private static void Validate(string source, LatexParseOptions options) {
        if (options.MaximumInputLength.HasValue && options.MaximumInputLength.Value < 0) throw new ArgumentOutOfRangeException(nameof(options));
        if (options.MaximumInputLength.HasValue && source.Length > options.MaximumInputLength.Value) throw new ArgumentException("LaTeX source exceeds MaximumInputLength.", nameof(source));
        if (options.MaximumTokenCount < 1) throw new ArgumentOutOfRangeException(nameof(options), "MaximumTokenCount must be positive.");
        if (options.MaximumNestingDepth < 1) throw new ArgumentOutOfRangeException(nameof(options), "MaximumNestingDepth must be positive.");
        if (options.MaximumExpansionDepth < 1) throw new ArgumentOutOfRangeException(nameof(options), "MaximumExpansionDepth must be positive.");
        if (options.MaximumExpansionLength < 1) throw new ArgumentOutOfRangeException(nameof(options), "MaximumExpansionLength must be positive.");
    }
}
