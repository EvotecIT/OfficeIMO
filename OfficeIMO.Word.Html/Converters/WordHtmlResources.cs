namespace OfficeIMO.Word.Html {
    internal static class WordHtmlResources {
        // Default CSS for documents exported without a named document-shell profile. The --office-* tokens come
        // from OfficeTheme.Default; the rules follow the visual design (horizontal table rules, tabular numbers,
        // text-coloured headings, monospace stacks ending in a generic family).
        private const string DefaultRules = "body{background-color:var(--office-bg);color:var(--office-text);margin:0;padding:18pt;font-family:var(--office-font-body);font-size:var(--office-size-body);line-height:1.4}"
            + "h1,h2,h3,h4,h5,h6{font-family:var(--office-font-heading);color:var(--office-text);font-weight:700;line-height:1.25;margin:1.2em 0 .4em}"
            + "h1{font-size:var(--office-size-h1)}h2{font-size:var(--office-size-h2)}h3{font-size:var(--office-size-h3)}h4{font-size:var(--office-size-h4)}h5{font-size:var(--office-size-h5)}h6{font-size:var(--office-size-h6)}"
            + "p{margin:0 0 var(--office-paragraph-after)}a{color:var(--office-link)}img{max-width:100%;height:auto}"
            + "table{width:100%;border-collapse:collapse;margin:0 0 12pt;font-variant-numeric:tabular-nums}"
            + "th,td{border-bottom:var(--office-table-border-width) solid var(--office-table-border);padding:var(--office-table-cell-padding-y) var(--office-table-cell-padding-x);vertical-align:top}"
            + "thead th{background-color:var(--office-table-header-fill);color:var(--office-table-header-text);font-weight:700}"
            + "tbody tr:nth-child(2n){background-color:var(--office-table-band-fill)}"
            + "code{font-family:var(--office-font-mono);font-size:.9em;background-color:var(--office-surface);padding:0 .25em;border-radius:2px}"
            + "pre{font-family:var(--office-font-mono);font-size:.9em;background-color:var(--office-surface);border:1px solid var(--office-rule);padding:8pt 10pt;border-radius:3pt;white-space:pre-wrap}"
            + "pre code{background:none;padding:0;font-size:inherit}"
            + "blockquote{color:var(--office-muted);border-left:3pt solid var(--office-rule);padding:0 12pt;margin:0 0 var(--office-paragraph-after)}"
            + "nav.toc{margin:0 0 16pt}nav.toc .toc-title{font-family:var(--office-font-heading);font-weight:700;font-size:var(--office-size-h2);margin:0 0 6pt}"
            + "nav.toc ul{list-style:none;padding-left:0;margin:0}nav.toc ul ul{padding-left:var(--office-list-indent)}nav.toc li{margin:2pt 0}nav.toc a{color:inherit;text-decoration:none}";

        // Rules carry literal values so HTML to Word, which does not evaluate custom properties, reads them back;
        // the tokens stay declared on :root for pages that restyle the output.
        private static readonly string DefaultCss =
            global::OfficeIMO.Internal.OfficeThemeCssTokens.Build(OfficeTheme.Default, ":root").Trim()
            + global::OfficeIMO.Internal.OfficeThemeCssTokens.ResolveReferences(DefaultRules, OfficeTheme.Default);

        internal static string GetDefaultCss(OfficeIMO.Drawing.OfficeVisualThemeKind theme, bool useSharedDocumentShell) =>
            (useSharedDocumentShell
                ? OfficeIMO.Html.OfficeHtmlDocumentShell.GetThemeCss(theme)
                : DefaultCss)
            + "li p{margin:0}";
    }
}
