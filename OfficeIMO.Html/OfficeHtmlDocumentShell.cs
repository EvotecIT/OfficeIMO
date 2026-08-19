namespace OfficeIMO.Html;

/// <summary>
/// Shared document shell and theme CSS for OfficeIMO-generated HTML adapters.
/// </summary>
public static class OfficeHtmlDocumentShell {
    private const string StyleResourceName = "OfficeIMO.Html.Assets.OfficeHtmlDocumentShell.css";
    private static readonly Lazy<string> SharedStyles = new Lazy<string>(LoadSharedStyles, LazyThreadSafetyMode.ExecutionAndPublication);

    /// <summary>
    /// Wraps a body fragment in a complete HTML document using the shared OfficeIMO shell.
    /// </summary>
    public static string WrapBody(string bodyHtml, OfficeHtmlDocumentOptions? options = null) {
        options ??= new OfficeHtmlDocumentOptions();
        options.Validate();
        if (!options.EmitDocumentShell) return bodyHtml ?? string.Empty;
        string nl = options.NewLine;
        var builder = new StringBuilder();
        builder.Append("<!doctype html>").Append(nl);
        builder.Append("<html lang=\"").Append(OfficeHtmlText.EscapeAttribute(options.Language ?? "en")).Append("\">").Append(nl);
        builder.Append("<head>").Append(nl);
        builder.Append("<meta charset=\"utf-8\">").Append(nl);
        builder.Append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">").Append(nl);
        builder.Append("<title>").Append(OfficeHtmlText.Escape(string.IsNullOrWhiteSpace(options.Title) ? "OfficeIMO HTML" : options.Title!)).Append("</title>").Append(nl);
        if (options.IncludeDefaultStyles) {
            builder.Append("<style>").Append(nl);
            builder.Append(GetThemeCss(options.Theme, nl)).Append(nl);
            builder.Append("</style>").Append(nl);
        }

        builder.Append("</head>").Append(nl);
        builder.Append("<body class=\"").Append(OfficeHtmlText.EscapeAttribute(options.BodyClass)).Append("\">").Append(nl);
        builder.Append(bodyHtml ?? string.Empty);
        if (!builder.ToString().EndsWith(nl, StringComparison.Ordinal)) {
            builder.Append(nl);
        }

        builder.Append("</body>").Append(nl);
        builder.Append("</html>").Append(nl);
        return builder.ToString();
    }

    /// <summary>
    /// Gets the shared CSS for an OfficeIMO HTML document theme.
    /// </summary>
    public static string GetThemeCss(OfficeVisualThemeKind theme, string newLine = "\n") {
        if (!Enum.IsDefined(typeof(OfficeVisualThemeKind), theme)) throw new ArgumentOutOfRangeException(nameof(theme));
        string nl = string.IsNullOrEmpty(newLine) ? "\n" : newLine;
        OfficeHtmlThemePalette palette = OfficeHtmlThemePalette.Create(theme);
        var builder = new StringBuilder();
        builder.Append(":root{")
            .Append("--officeimo-accent:").Append(palette.Accent).Append(';')
            .Append("--officeimo-accent-dark:").Append(palette.AccentDark).Append(';')
            .Append("--officeimo-accent-soft:").Append(palette.AccentSoft).Append(';')
            .Append("--officeimo-page-background:").Append(palette.PageBackground).Append(';')
            .Append("--officeimo-surface:").Append(palette.Surface).Append(';')
            .Append("--officeimo-panel:").Append(palette.Panel).Append(';')
            .Append("--officeimo-border:").Append(palette.Border).Append(';')
            .Append("--officeimo-border-strong:").Append(palette.BorderStrong).Append(';')
            .Append("--officeimo-text:").Append(palette.Text).Append(';')
            .Append("--officeimo-heading:").Append(palette.Heading).Append(';')
            .Append("--officeimo-muted:").Append(palette.Muted).Append(';')
            .Append("--officeimo-table-header:").Append(palette.TableHeader).Append(';')
            .Append("--officeimo-table-stripe:").Append(palette.TableStripe).Append(';')
            .Append("--officeimo-code-background:").Append(palette.CodeBackground).Append(';')
            .Append("--officeimo-control-disabled:").Append(palette.ControlDisabled).Append(';')
            .Append("--officeimo-warning:").Append(palette.Warning).Append(';')
            .Append("--officeimo-warning-soft:").Append(palette.WarningSoft).Append(';')
            .Append("--officeimo-font-body:").Append(palette.FontBody).Append(';')
            .Append("--officeimo-font-heading:").Append(palette.FontHeading).Append(';')
            .Append("--officeimo-font-mono:").Append(palette.FontMono).Append(';')
            .Append("--officeimo-font-size:").Append(palette.FontSize).Append(";}")
            .Append(nl)
            .Append(NormalizeNewLines(SharedStyles.Value, nl));
        return builder.ToString();
    }

    /// <summary>
    /// Combines required adapter classes with caller-supplied classes, removing duplicate tokens
    /// while preserving first-seen order.
    /// </summary>
    public static string MergeBodyClasses(params string?[] classLists) {
        if (classLists == null) throw new ArgumentNullException(nameof(classLists));
        var classes = new List<string>();
        var seen = new HashSet<string>(StringComparer.Ordinal);
        foreach (string? classList in classLists) {
            if (string.IsNullOrWhiteSpace(classList)) continue;
            foreach (string token in classList!.Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries)) {
                if (seen.Add(token)) classes.Add(token);
            }
        }

        return string.Join(" ", classes);
    }

    private static string LoadSharedStyles() {
        using Stream stream = typeof(OfficeHtmlDocumentShell).Assembly.GetManifestResourceStream(StyleResourceName)
            ?? throw new InvalidOperationException("Embedded Office HTML shell stylesheet is missing.");
        using var reader = new StreamReader(stream, Encoding.UTF8, true);
        return reader.ReadToEnd();
    }

    private static string NormalizeNewLines(string value, string newLine) =>
        value.Replace("\r\n", "\n").Replace('\r', '\n').Replace("\n", newLine);
}
