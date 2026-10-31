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

    private PdfRedactionSearchCriteria(string[] literals, Regex[] expressions, bool coversAllPages, IReadOnlyCollection<string> affectedFields, IReadOnlyCollection<int> pageNumbers) {
        _literals = literals;
        _expressions = expressions;
        CoversAllPages = coversAllPages;
        AffectedFields = affectedFields;
        PageNumbers = pageNumbers;
    }

    /// <summary>True when the search wasn't limited to pages, so content tied to no page is in scope.</summary>
    internal bool CoversAllPages { get; }
    internal IReadOnlyCollection<int> PageNumbers { get; }

    /// <summary>Names of the fields this search clears. A name survives the renumbering of a rewrite, so identity can skip them.</summary>
    internal IReadOnlyCollection<string> AffectedFields { get; }

    internal PdfRedactionSearchCriteria WithAffectedFields(IEnumerable<string> fields) =>
        new PdfRedactionSearchCriteria(_literals, _expressions, CoversAllPages, new HashSet<string>(fields, StringComparer.Ordinal), PageNumbers);

    internal static PdfRedactionSearchCriteria? Create(IEnumerable<string> literals, IEnumerable<Regex> expressions, bool coversAllPages, IReadOnlyCollection<int>? pageNumbers = null) {
        string[] literalCriteria = literals.ToArray();
        Regex[] expressionCriteria = expressions.ToArray();
        return literalCriteria.Length == 0 && expressionCriteria.Length == 0
            ? null : new PdfRedactionSearchCriteria(literalCriteria, expressionCriteria, coversAllPages, Array.Empty<string>(), pageNumbers ?? Array.Empty<int>());
    }

    internal bool Matches(string text) =>
        _literals.Any(literal => PdfTextSearchNormalization.Contains(text, literal, StringComparison.OrdinalIgnoreCase)) ||
        _expressions.Any(expression => expression.IsMatch(text));
}
