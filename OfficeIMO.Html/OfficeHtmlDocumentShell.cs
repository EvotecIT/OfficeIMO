using System.Globalization;

namespace OfficeIMO.Html;

/// <summary>
/// Shared document shell and theme CSS for OfficeIMO-generated HTML adapters.
/// </summary>
public static class OfficeHtmlDocumentShell {
    /// <summary>
    /// Wraps a body fragment in a complete HTML document using the shared OfficeIMO shell.
    /// </summary>
    public static string WrapBody(string bodyHtml, OfficeHtmlDocumentOptions? options = null) {
        options ??= new OfficeHtmlDocumentOptions();
        string nl = string.IsNullOrEmpty(options.NewLine) ? "\n" : options.NewLine;
        var builder = new StringBuilder();
        builder.Append("<!doctype html>").Append(nl);
        builder.Append("<html lang=\"en\">").Append(nl);
        builder.Append("<head>").Append(nl);
        builder.Append("<meta charset=\"utf-8\">").Append(nl);
        builder.Append("<meta name=\"viewport\" content=\"width=device-width, initial-scale=1\">").Append(nl);
        builder.Append("<title>").Append(OfficeHtmlText.Escape(options.Title)).Append("</title>").Append(nl);
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
    public static string GetThemeCss(OfficeHtmlDocumentThemeKind theme, string newLine = "\n") {
        string accent;
        string accentDark;
        string surface;
        string border;
        string text;
        string muted;
        double fontSize;

        switch (theme) {
            case OfficeHtmlDocumentThemeKind.Compact:
                accent = "#0F766E";
                accentDark = "#115E59";
                surface = "#F8FAFC";
                border = "#CBD5E1";
                text = "#111827";
                muted = "#475569";
                fontSize = 13D;
                break;
            case OfficeHtmlDocumentThemeKind.Report:
                accent = "#1D4ED8";
                accentDark = "#1E3A8A";
                surface = "#F9FAFB";
                border = "#CBD5E1";
                text = "#111827";
                muted = "#475569";
                fontSize = 14D;
                break;
            case OfficeHtmlDocumentThemeKind.Technical:
                accent = "#047857";
                accentDark = "#065F46";
                surface = "#F8FAFC";
                border = "#CBD5E1";
                text = "#0F172A";
                muted = "#475569";
                fontSize = 13.5D;
                break;
            default:
                accent = "#2563EB";
                accentDark = "#1E40AF";
                surface = "#FFFFFF";
                border = "#D1D5DB";
                text = "#111827";
                muted = "#4B5563";
                fontSize = 14D;
                break;
        }

        string size = fontSize.ToString("0.##", CultureInfo.InvariantCulture);
        var builder = new StringBuilder();
        builder.Append(":root{--officeimo-accent:").Append(accent).Append(";--officeimo-accent-dark:").Append(accentDark).Append(";--officeimo-surface:").Append(surface).Append(";--officeimo-border:").Append(border).Append(";--officeimo-text:").Append(text).Append(";--officeimo-muted:").Append(muted).Append(";}").Append(newLine);
        builder.Append("body.officeimo-html{margin:0;background:#F3F4F6;color:var(--officeimo-text);font-family:\"Segoe UI\",Arial,sans-serif;font-size:").Append(size).Append("px;line-height:1.5;}").Append(newLine);
        builder.Append("main.officeimo-document{max-width:1120px;margin:0 auto;padding:28px 20px 40px;}").Append(newLine);
        builder.Append("h1,h2,h3{color:var(--officeimo-accent-dark);font-weight:650;line-height:1.2;margin:0 0 12px;}").Append(newLine);
        builder.Append("h1{font-size:28px;margin-bottom:18px;}h2{font-size:21px;margin-top:24px;}h3{font-size:16px;margin-top:18px;}").Append(newLine);
        builder.Append("p{margin:0 0 12px;}a{color:var(--officeimo-accent);}").Append(newLine);
        builder.Append(".officeimo-panel,.officeimo-sheet,.officeimo-slide{background:var(--officeimo-surface);border:1px solid var(--officeimo-border);border-radius:8px;padding:18px;margin:0 0 18px;box-shadow:0 1px 2px rgba(15,23,42,.06);}").Append(newLine);
        builder.Append(".officeimo-muted,.officeimo-diagnostic{color:var(--officeimo-muted);font-size:12px;}").Append(newLine);
        builder.Append(".officeimo-feature{border-top:1px solid var(--officeimo-border);margin-top:14px;padding-top:12px;}.officeimo-feature h3{margin-top:0;}.officeimo-feature-list{display:grid;grid-template-columns:repeat(auto-fit,minmax(220px,1fr));gap:10px;margin:8px 0 0;padding:0;list-style:none;}.officeimo-feature-item{border:1px solid var(--officeimo-border);border-radius:6px;background:#fff;padding:10px;}.officeimo-feature-label{display:block;font-weight:650;color:var(--officeimo-accent-dark);}.officeimo-feature-meta{color:var(--officeimo-muted);font-size:12px;}.officeimo-inline-image{display:block;max-width:100%;height:auto;margin-top:8px;border:1px solid var(--officeimo-border);border-radius:4px;}").Append(newLine);
        builder.Append("table.officeimo-table{border-collapse:collapse;width:100%;margin:10px 0 16px;background:#fff;}").Append(newLine);
        builder.Append(".officeimo-table th,.officeimo-table td{border:1px solid var(--officeimo-border);padding:6px 8px;vertical-align:top;text-align:left;}").Append(newLine);
        builder.Append(".officeimo-table th{background:#EFF6FF;color:#1E3A8A;font-weight:650;}").Append(newLine);
        builder.Append(".officeimo-visual-page{overflow:auto;background:#fff;border:1px solid var(--officeimo-border);border-radius:8px;margin:12px 0 20px;padding:12px;}").Append(newLine);
        builder.Append(".officeimo-slide-canvas{position:relative;background:#fff;border:1px solid var(--officeimo-border);box-shadow:0 1px 2px rgba(15,23,42,.08);overflow:hidden;}").Append(newLine);
        builder.Append(".officeimo-shape{position:absolute;box-sizing:border-box;border:1px solid rgba(37,99,235,.20);padding:6px;overflow:hidden;background:rgba(255,255,255,.82);}").Append(newLine);
        builder.Append(".officeimo-shape-table{padding:0;background:#fff;}.officeimo-shape-picture{padding:0;background:transparent;border-color:transparent;}.officeimo-shape-picture img{width:100%;height:100%;object-fit:contain;display:block;}.officeimo-shape-chart{padding:0;background:#fff;}.officeimo-chart-rendered,.officeimo-chart-rendered svg{display:block;width:100%;height:100%;}.officeimo-shape-placeholder{display:flex;align-items:center;justify-content:center;color:var(--officeimo-muted);font-size:12px;background:#F8FAFC;width:100%;height:100%;text-align:center;}.officeimo-chart-placeholder{align-items:flex-start;justify-content:flex-start;padding:10px;background:#F8FAFC;}.officeimo-chart-bars{display:flex;align-items:flex-end;gap:4px;height:64px;margin-top:8px;border-bottom:1px solid var(--officeimo-border);}.officeimo-chart-bars span{display:block;min-width:10px;background:var(--officeimo-accent);border-radius:3px 3px 0 0;}").Append(newLine);
        builder.Append("pre.officeimo-source-markdown{white-space:pre-wrap;background:#F8FAFC;border:1px solid var(--officeimo-border);border-radius:6px;padding:10px;overflow:auto;}").Append(newLine);
        return builder.ToString();
    }
}
