using System.Text.RegularExpressions;

namespace OfficeIMO.Workflows;

public sealed partial class BookProject {
    // Lexical checks only. Assignment, namespace-specific rules and resolver availability
    // belong to the identifier authority and are never inferred or fetched during export.
    private static string RequireOnixCollectionDoi(string value) {
        int slash = value.IndexOf('/');
        if (!value.StartsWith("10.", StringComparison.Ordinal) || slash < 4 || slash == value.Length - 1 ||
            value.Any(char.IsControl) || value.Substring(3, slash - 3).Split('.').Any(part =>
                part.Length == 0 || part.Any(character => character < '0' || character > '9')))
            throw new ArgumentException("Supply a bare DOI with a numeric registrant prefix and nonempty suffix, without a resolver URL.", nameof(value));
        return value;
    }

    private static string RequireOnixCollectionUrn(string value) {
        // RFC 8141 assigned-name and optional resolution, query and fragment components.
        const string pchar = "(?:[A-Za-z0-9._~!$&'()*+,;=:@-]|%[0-9A-Fa-f]{2})";
        const string tail = "(?:" + pchar + "|[/?])*";
        const string pattern = "\\A(?i:urn):[A-Za-z0-9][A-Za-z0-9-]{0,30}[A-Za-z0-9]:" + pchar +
            "(?:" + pchar + "|/)*(?:\\?\\+" + pchar + tail + ")?(?:\\?=" + pchar + tail + ")?(?:#" + tail + ")?\\z";
        int fragment = value.IndexOf('#');
        string withoutFragment = fragment < 0 ? value : value.Substring(0, fragment);
        int query = withoutFragment.IndexOf("?=", StringComparison.Ordinal);
        bool queryValid = query < 0 || Regex.IsMatch(withoutFragment.Substring(query + 2),
            "\\A" + pchar + tail + "\\z", RegexOptions.CultureInvariant | RegexOptions.NonBacktracking);
        if (!queryValid || !Regex.IsMatch(value, pattern, RegexOptions.CultureInvariant | RegexOptions.NonBacktracking))
            throw new ArgumentException("Supply a full URN using RFC 8141 syntax.", nameof(value));
        return value;
    }

    private static string RequireOnixMagazineIdentifier(string value) {
        if (value.Length != 5 || value.Any(character => character < '0' || character > '9'))
            throw new ArgumentException("A Japanese magazine identifier contains exactly five ASCII digits, without an issue extension.", nameof(value));
        return value;
    }

    private static string RequireOnixCollectionArk(string value) {
        if (!Uri.TryCreate(value, UriKind.Absolute, out Uri? uri) || !uri.IsWellFormedOriginalString() ||
            (uri.Scheme != Uri.UriSchemeHttp && uri.Scheme != Uri.UriSchemeHttps) || string.IsNullOrEmpty(uri.Host))
            throw new ArgumentException("Supply an ARK including its HTTP or HTTPS resolver URL.", nameof(value));
        // Inspect the supplied path: Uri.AbsolutePath can normalize dot segments and
        // accidentally turn a malformed authority into an apparently valid one.
        int authorityEnd = value.IndexOfAny(['/', '?', '#'], value.IndexOf("://", StringComparison.Ordinal) + 3);
        string path = authorityEnd < 0 || value[authorityEnd] != '/' ? string.Empty : value.Substring(authorityEnd).Split('?', '#')[0];
        int marker = path.IndexOf("/ark:/", StringComparison.Ordinal);
        string name = marker < 0 ? string.Empty : path.Substring(marker + 6);
        int slash = name.IndexOf('/');
        if (slash < 1 || slash == name.Length - 1 || name.Substring(0, slash).Any(character => character < '0' || character > '9'))
            throw new ArgumentException("The ARK resolver URL must contain /ark:/ followed by a numeric authority and nonempty name.", nameof(value));
        return value;
    }
}
