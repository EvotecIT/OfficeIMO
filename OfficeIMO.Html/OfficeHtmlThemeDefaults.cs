namespace OfficeIMO.Html;

/// <summary>Projects shared theme tokens into document defaults without adding a presentation shell.</summary>
internal static class OfficeHtmlThemeDefaults {
    private static readonly Lazy<string> Styles = new(LoadStyles, LazyThreadSafetyMode.ExecutionAndPublication);

    internal static string Build(OfficeTheme theme, string scope) =>
        global::OfficeIMO.Internal.OfficeThemeCssTokens.ResolveReferences(
            Styles.Value.Replace("__SCOPE__", scope), theme);

    private static string LoadStyles() {
        using Stream stream = typeof(OfficeHtmlThemeDefaults).Assembly.GetManifestResourceStream("OfficeIMO.Html.Assets.OfficeThemeDefaults.css")
            ?? throw new InvalidOperationException("Embedded theme defaults stylesheet is missing.");
        using var reader = new StreamReader(stream, Encoding.UTF8, true);
        return reader.ReadToEnd();
    }
}
