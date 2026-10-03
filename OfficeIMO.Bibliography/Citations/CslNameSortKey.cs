namespace OfficeIMO.Bibliography;

/// <summary>Separates name-key priorities that must not depend on display punctuation.</summary>
internal static class CslNameSortKey {
    // These separators are used only in transient sort keys, never rendered output.
    internal const string ListDelimiter = "\u001e";
    private const string PartDelimiter = "\u001f";
    private static readonly string[] InstitutionArticles = { "a", "an", "the" };

    internal static CslText Personal(string family, string particles, string given, string suffix) {
        if (family.Length == 0 && particles.Length == 0 && given.Length == 0 && suffix.Length == 0) return CslText.Empty;
        return CslText.Literal(string.Join(PartDelimiter, family, particles, given, suffix).TrimEnd('\u001f'));
    }

    internal static bool IsSeparator(string value) => value.Length == 1 && (value[0] == '\u001e' || value[0] == '\u001f');

    internal static string StripInstitutionArticle(string value) {
        foreach (string article in InstitutionArticles) {
            if (!value.StartsWith(article, StringComparison.OrdinalIgnoreCase) || value.Length == article.Length || !char.IsWhiteSpace(value[article.Length])) continue;
            int start = article.Length;
            while (start < value.Length && char.IsWhiteSpace(value[start])) start++;
            return value.Substring(start);
        }
        return value;
    }
}
