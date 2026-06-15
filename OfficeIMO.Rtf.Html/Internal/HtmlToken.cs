namespace OfficeIMO.Rtf.Html;

internal sealed class HtmlToken {
    internal HtmlToken(HtmlTokenKind kind, string value, IReadOnlyDictionary<string, string>? attributes = null, bool selfClosing = false) {
        Kind = kind;
        Value = value;
        Attributes = attributes ?? EmptyAttributes;
        SelfClosing = selfClosing;
    }

    internal HtmlTokenKind Kind { get; }

    internal string Value { get; }

    internal IReadOnlyDictionary<string, string> Attributes { get; }

    internal bool SelfClosing { get; }

    private static readonly IReadOnlyDictionary<string, string> EmptyAttributes = new Dictionary<string, string>();
}
