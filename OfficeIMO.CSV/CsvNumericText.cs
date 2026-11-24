using System.Text.RegularExpressions;

namespace OfficeIMO.CSV;

internal sealed class CsvNumericText
{
    private static readonly Regex Number = new(
        @"\A-?(?:0|[1-9][0-9]*)(?:\.[0-9]+)?(?:[eE][+-]?[0-9]+)?\z",
        RegexOptions.CultureInvariant);
    private readonly string _text;

    // Preserve the exact invariant spelling of a numeric scalar without treating it as a formula.
    internal CsvNumericText(string text)
    {
        if (!Number.IsMatch(text)) throw new ArgumentException("Expected an invariant number.", nameof(text));
        _text = text;
    }

    public override string ToString() => _text;
}
