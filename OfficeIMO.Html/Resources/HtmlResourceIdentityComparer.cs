namespace OfficeIMO.Html;

/// <summary>
/// Compares resource identities using URI component semantics: scheme and host are
/// case-insensitive, while path, query, user information, and relative references
/// remain case-sensitive. Content identifiers are case-insensitive.
/// </summary>
internal sealed class HtmlResourceIdentityComparer : IEqualityComparer<string> {
    internal static readonly HtmlResourceIdentityComparer Instance = new HtmlResourceIdentityComparer();

    private HtmlResourceIdentityComparer() { }

    public bool Equals(string? left, string? right) {
        if (ReferenceEquals(left, right)) return true;
        if (left == null || right == null) return false;
        if (Uri.TryCreate(left, UriKind.Absolute, out Uri? leftUri)
            && Uri.TryCreate(right, UriKind.Absolute, out Uri? rightUri)) {
            return Equals(leftUri, rightUri);
        }
        return string.Equals(left, right, StringComparison.Ordinal);
    }

    public int GetHashCode(string value) {
        if (!Uri.TryCreate(value, UriKind.Absolute, out Uri? uri)) {
            return StringComparer.Ordinal.GetHashCode(value);
        }
        if (uri.Scheme.Equals("cid", StringComparison.OrdinalIgnoreCase)) {
            return StringComparer.OrdinalIgnoreCase.GetHashCode(uri.AbsoluteUri);
        }
        unchecked {
            int hash = StringComparer.OrdinalIgnoreCase.GetHashCode(uri.Scheme);
            hash = (hash * 397) ^ StringComparer.OrdinalIgnoreCase.GetHashCode(uri.IdnHost);
            hash = (hash * 397) ^ uri.Port;
            hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(uri.UserInfo);
            hash = (hash * 397) ^ StringComparer.Ordinal.GetHashCode(
                uri.GetComponents(UriComponents.PathAndQuery, UriFormat.UriEscaped));
            return hash;
        }
    }

    internal static bool Equals(Uri left, Uri right) {
        if (!left.Scheme.Equals(right.Scheme, StringComparison.OrdinalIgnoreCase)) return false;
        if (left.Scheme.Equals("cid", StringComparison.OrdinalIgnoreCase)) {
            return left.AbsoluteUri.Equals(right.AbsoluteUri, StringComparison.OrdinalIgnoreCase);
        }
        return left.IdnHost.Equals(right.IdnHost, StringComparison.OrdinalIgnoreCase)
            && left.Port == right.Port
            && left.UserInfo.Equals(right.UserInfo, StringComparison.Ordinal)
            && left.GetComponents(UriComponents.PathAndQuery, UriFormat.UriEscaped).Equals(
                right.GetComponents(UriComponents.PathAndQuery, UriFormat.UriEscaped),
                StringComparison.Ordinal);
    }
}

internal sealed class HtmlResourceSeenKeyComparer : IEqualityComparer<string> {
    internal static readonly HtmlResourceSeenKeyComparer Instance = new HtmlResourceSeenKeyComparer();

    private HtmlResourceSeenKeyComparer() { }

    public bool Equals(string? left, string? right) {
        if (ReferenceEquals(left, right)) return true;
        if (left == null || right == null) return false;
        int leftSeparator = left.IndexOf('\n');
        int rightSeparator = right.IndexOf('\n');
        if (leftSeparator < 0 || rightSeparator < 0) return string.Equals(left, right, StringComparison.Ordinal);
        return leftSeparator == rightSeparator
            && string.CompareOrdinal(left, 0, right, 0, leftSeparator) == 0
            && HtmlResourceIdentityComparer.Instance.Equals(left.Substring(leftSeparator + 1), right.Substring(rightSeparator + 1));
    }

    public int GetHashCode(string value) {
        int separator = value.IndexOf('\n');
        if (separator < 0) return StringComparer.Ordinal.GetHashCode(value);
        unchecked {
            int hash = StringComparer.Ordinal.GetHashCode(value.Substring(0, separator));
            return (hash * 397) ^ HtmlResourceIdentityComparer.Instance.GetHashCode(value.Substring(separator + 1));
        }
    }
}
