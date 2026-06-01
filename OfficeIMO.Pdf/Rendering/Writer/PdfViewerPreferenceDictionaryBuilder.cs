namespace OfficeIMO.Pdf;

internal static class PdfViewerPreferenceDictionaryBuilder {
    internal static string BuildGeneratedViewerPreferencesDictionary(PdfViewerPreferencesOptions preferences) {
        Guard.NotNull(preferences, nameof(preferences));
        if (!preferences.HasAny) {
            throw new ArgumentException("At least one PDF viewer preference must be configured.", nameof(preferences));
        }

        var sb = new StringBuilder();
        sb.Append("<<");
        AppendBooleanEntry(sb, "HideToolbar", preferences.HideToolbar);
        AppendBooleanEntry(sb, "HideMenubar", preferences.HideMenubar);
        AppendBooleanEntry(sb, "HideWindowUI", preferences.HideWindowUI);
        AppendBooleanEntry(sb, "FitWindow", preferences.FitWindow);
        AppendBooleanEntry(sb, "CenterWindow", preferences.CenterWindow);
        AppendBooleanEntry(sb, "DisplayDocTitle", preferences.DisplayDocTitle);
        sb.Append(" >>\n");
        return sb.ToString();
    }

    private static void AppendBooleanEntry(StringBuilder sb, string key, bool? value) {
        if (!value.HasValue) {
            return;
        }

        sb.Append(" /")
            .Append(PdfSyntaxEscaper.Name(key))
            .Append(value.Value ? " true" : " false");
    }
}
