using System.Text.RegularExpressions;

namespace OfficeIMO.Word;

public partial class WordSection {
    /// <summary>Finds literal text in this section's body, tables, and resolved headers and footers using the native search groups.</summary>
    /// <param name="text">Text to find; a match can span runs and, unless WholeWords is selected, paragraphs in one part.</param>
    /// <param name="options">Comparison settings. Null ignores case and permits matches inside words.</param>
    /// <returns>Paragraphs containing matches, grouped by part, and the total match count.</returns>
    public WordSearchResult Find(string text, OfficeTextSearchOptions? options = null) => _document.FindInScope(text, options, this);

    /// <summary>Finds regular-expression matches within individual paragraphs in this section's body, tables, and resolved headers and footers.</summary>
    /// <param name="pattern">Pattern whose options control comparison and boundaries.</param>
    /// <returns>Paragraphs containing matches, grouped by part, and the total match count.</returns>
    public WordSearchResult Find(Regex pattern) => _document.FindInScope(pattern, this);
}
