namespace OfficeIMO.OpenDocument;

/// <summary>Normalizes internal package references without resolving or fetching external content.</summary>
internal static class OdfPackagePath {
    /// <summary>Removes query/fragment suffixes, leading relative markers, and URI escaping from an internal href.</summary>
    internal static string NormalizeHref(string href) {
        string value = href;
        int fragment = value.IndexOf('#');
        if (fragment >= 0) value = value.Substring(0, fragment);
        int query = value.IndexOf('?');
        if (query >= 0) value = value.Substring(0, query);
        while (value.StartsWith("./", StringComparison.Ordinal)) value = value.Substring(2);
        try { return Uri.UnescapeDataString(value); } catch (UriFormatException) { return value; }
    }
}
