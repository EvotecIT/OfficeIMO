namespace OfficeIMO.AsciiDoc;

/// <summary>Shared constrained-marker boundaries for native parsing and semantic writers.</summary>
internal static class AsciiDocInlineFormattingRules {
    internal static bool CanOpen(char? before, char? first) =>
        first.HasValue && !char.IsWhiteSpace(first.Value) &&
        (!before.HasValue || (!IsWord(before.Value) && before != ':' && before != ';' && before != '}'));

    internal static bool CanClose(char? last, char? after) =>
        last.HasValue && !char.IsWhiteSpace(last.Value) &&
        (!after.HasValue || (!IsWord(after.Value) && after != ':' && after != ';' && after != '{'));

    private static bool IsWord(char value) => char.IsLetterOrDigit(value) || value == '_';
}
