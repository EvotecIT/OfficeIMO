namespace OfficeIMO.Html;

/// <summary>A source-preserving operand change or a bounded set of exact attribute matches.</summary>
internal sealed class HtmlCssAttributeSelectorEdit {
    internal const int MaximumAlternatives = 256;
    internal const int MaximumSelectorLength = 65536;
    internal string? Value { get; }
    internal IReadOnlyList<string>? ExactValues { get; }

    private HtmlCssAttributeSelectorEdit(string? value, IReadOnlyList<string>? exactValues) {
        Value = value; ExactValues = exactValues;
    }

    internal static HtmlCssAttributeSelectorEdit Operand(string value) => new HtmlCssAttributeSelectorEdit(value, null);
    internal static HtmlCssAttributeSelectorEdit Exact(IReadOnlyList<string> values) => new HtmlCssAttributeSelectorEdit(null, values);
}
