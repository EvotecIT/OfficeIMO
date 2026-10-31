using System.Text.RegularExpressions;

namespace OfficeIMO.Pdf;

/// <summary>
/// The literal and regular-expression criteria behind a search-derived area. They travel with the area
/// so a plan rebuilt from selected areas can still clear and check copies of the searched text held
/// outside the page content. Case is ignored here because keeping a copy is the worse failure.
/// </summary>
internal sealed class PdfRedactionSearchCriteria {
    private readonly string[] _literals;
    private readonly Regex[] _expressions;

    private PdfRedactionSearchCriteria(string[] literals, Regex[] expressions) {
        _literals = literals;
        _expressions = expressions;
    }

    internal static PdfRedactionSearchCriteria? Create(IEnumerable<string> literals, IEnumerable<Regex> expressions) {
        string[] literalCriteria = literals.ToArray();
        Regex[] expressionCriteria = expressions.ToArray();
        return literalCriteria.Length == 0 && expressionCriteria.Length == 0
            ? null : new PdfRedactionSearchCriteria(literalCriteria, expressionCriteria);
    }

    internal bool Matches(string text) =>
        _literals.Any(literal => PdfTextSearchNormalization.Contains(text, literal, StringComparison.OrdinalIgnoreCase)) ||
        _expressions.Any(expression => expression.IsMatch(text));
}
