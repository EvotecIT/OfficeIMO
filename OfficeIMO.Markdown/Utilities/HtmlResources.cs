namespace OfficeIMO.Markdown;

/// <summary>
/// Built-in CSS and tiny scripts for HTML rendering.
/// </summary>
    internal static class HtmlResources {
        internal static string GetStyleCss(MarkdownHtmlStyle style) => style switch {
            MarkdownHtmlStyle.Plain => string.Empty,
            MarkdownHtmlStyle.Clean => CleanThemeCss,
            MarkdownHtmlStyle.GithubLight => GithubLightCss,
            MarkdownHtmlStyle.GithubDark => GithubDarkCss,
            MarkdownHtmlStyle.GithubAuto => GithubAutoCss,
            MarkdownHtmlStyle.ChatLight => ChatLightCss,
            MarkdownHtmlStyle.ChatDark => ChatDarkCss,
            MarkdownHtmlStyle.ChatAuto => ChatAutoCss,
            MarkdownHtmlStyle.Word => WordCss,
            _ => CleanThemeCss
        };

    // The default style follows the 3.6 visual design: OfficeTheme tokens (--office-*) drive the Clean rules.
    private static readonly string CleanThemeCss =
        global::OfficeIMO.Internal.OfficeThemeCssTokens.Build(global::OfficeIMO.OfficeTheme.Default, ":root") + CleanCss;

    // Modern defaults from the shared theme tokens; literal fallbacks keep the rules usable without them.
    private const string CleanCss = @"
html,body { height: 100%; }
body { background: var(--office-bg, #ffffff); color: var(--office-text, #1b2433); margin: 0; }
article.markdown-body { max-width: 860px; margin: 2rem auto; padding: 0 1rem; line-height: 1.5; font-family: var(--office-font-body, Aptos, Calibri, Carlito, ""Helvetica Neue"", Arial, sans-serif); font-size: var(--office-size-body, 11pt); color: inherit; }
article.markdown-body h1,article.markdown-body h2,article.markdown-body h3,article.markdown-body h4,article.markdown-body h5,article.markdown-body h6 { font-family: var(--office-font-heading, Aptos, Calibri, Carlito, ""Helvetica Neue"", Arial, sans-serif); color: var(--office-text, #1b2433); font-weight: 700; line-height: 1.25; margin: 1.4em 0 .4em; }
article.markdown-body h1 { font-size: var(--office-size-h1, 20pt); margin-top: .6em; }
article.markdown-body h2 { font-size: var(--office-size-h2, 15pt); }
article.markdown-body h3 { font-size: var(--office-size-h3, 12.5pt); }
article.markdown-body h4 { font-size: var(--office-size-h4, 11pt); }
article.markdown-body h5 { font-size: var(--office-size-h5, 11pt); }
article.markdown-body h6 { font-size: var(--office-size-h6, 10pt); color: var(--office-muted, #5b6575); }
article.markdown-body p { margin: 0 0 var(--office-paragraph-after, 8pt); }
article.markdown-body a { color: var(--office-link, #3f64d7); text-decoration: underline; text-underline-offset: .15em; }
article.markdown-body a:hover { text-decoration-thickness: 2px; }
article.markdown-body ul, article.markdown-body ol { margin: 0 0 var(--office-paragraph-after, 8pt); padding-left: 1.6em; }
article.markdown-body li { margin: .15em 0; }
article.markdown-body li::marker { color: var(--office-accent-1, #3f64d7); }
article.markdown-body table { border-collapse: collapse; width: 100%; font-size: .95em; font-variant-numeric: tabular-nums; margin: 0 0 1em; }
article.markdown-body th, article.markdown-body td { padding: var(--office-table-cell-padding-y, 5pt) var(--office-table-cell-padding-x, 7pt); border-bottom: var(--office-table-border-width, .5pt) solid var(--office-table-border, #dce2ea); vertical-align: top; text-align: left; }
article.markdown-body thead th { background: var(--office-table-header-fill, #eaf0fc); color: var(--office-table-header-text, #1b2433); font-weight: 700; border-bottom: 1pt solid var(--office-accent-1, #3f64d7); }
article.markdown-body tbody tr:nth-child(2n) { background: var(--office-table-band-fill, #f6f8fb); }
article.markdown-body code { font-family: var(--office-font-mono, ui-monospace, Menlo, Consolas, monospace); font-size: .9em; background: var(--office-surface, #f6f8fb); border: 1px solid var(--office-rule, #dce2ea); padding: .05em .3em; border-radius: 3px; }
article.markdown-body pre { font-family: var(--office-font-mono, ui-monospace, Menlo, Consolas, monospace); font-size: .9em; line-height: 1.45; background: var(--office-surface, #f6f8fb); border: 1px solid var(--office-rule, #dce2ea); padding: 10px 12px; border-radius: 4px; overflow: auto; }
article.markdown-body pre code { background: none; border: 0; padding: 0; border-radius: 0; font-size: inherit; }
article.markdown-body td code { white-space: nowrap; }
article.markdown-body blockquote { color: var(--office-muted, #5b6575); border-left: 3px solid var(--office-rule, #dce2ea); padding: 0 1em; margin: 0 0 var(--office-paragraph-after, 8pt); }
article.markdown-body blockquote.callout { color: inherit; border-left-color: var(--office-accent-1, #3f64d7); background: var(--office-accent-1-tint, #ecf0fb); padding: .5em 1em; }
article.markdown-body blockquote.callout.warning { border-left-color: var(--office-warning, #865500); background: var(--office-warning-fill, #f0ebe0); }
article.markdown-body blockquote.callout.danger { border-left-color: var(--office-danger, #bf3b2c); background: var(--office-danger-fill, #f7e7e6); }
article.markdown-body hr { border: 0; border-top: 1px solid var(--office-rule, #dce2ea); margin: 1.2em 0; }
article.markdown-body img { max-width: 100%; height: auto; }
.anchor { margin-right: .25em; opacity: .6; text-decoration: none; }
/* data-theme overrides */
html[data-theme=dark] body { background: #0d1117; color: #c9d1d9; }
html[data-theme=light] body { background: #ffffff; color: #24292f; }
.theme-toggle { position: fixed; top: 12px; right: 12px; z-index: 9999; border: 1px solid rgba(27,31,36,.15); background: rgba(240,246,252,.9); color: inherit; border-radius: 6px; padding: 6px 8px; cursor: pointer; }
@media (prefers-color-scheme: dark) { .theme-toggle { border-color: rgba(240,246,252,.2); background: rgba(22,27,34,.85); } }
html[data-theme=dark] .theme-toggle { border-color: rgba(240,246,252,.2); background: rgba(22,27,34,.85); }
/* anchor + back-to-top UX */
.heading-anchor { margin-left: .4rem; opacity: 0; text-decoration: none; font-size: .9em; }
article.markdown-body h1:hover .heading-anchor,
article.markdown-body h2:hover .heading-anchor,
article.markdown-body h3:hover .heading-anchor,
article.markdown-body h4:hover .heading-anchor,
article.markdown-body h5:hover .heading-anchor,
article.markdown-body h6:hover .heading-anchor { opacity: .8; }
.back-to-top { margin: .35rem 0 .75rem 0; }
.back-to-top a { font-size: .85em; opacity: .8; }
/* Collapsible TOC details */
details.md-toc { margin: .5rem 0 1rem 0; }
details.md-toc > summary { cursor: pointer; font-weight: 600; margin-bottom: .5rem; }
/* Printed pages carry their number in the bottom margin */
@page { @bottom-right { content: ""Page "" counter(page) "" of "" counter(pages); font-family: var(--office-font-body, Aptos, Calibri, Carlito, ""Helvetica Neue"", Arial, sans-serif); font-size: 8pt; color: var(--office-muted, #5b6575); } }
";

    // Compact, chat-friendly styling. Intentionally avoids large page margins and keeps typography tighter.
    // Marker comment is used by tests.
    private const string ChatLightCss = @"
/* omd-chat */
html,body { height: 100%; }
body { background: transparent; color: #111827; margin: 0; }
article.markdown-body { max-width: 100%; margin: 0; padding: 0; line-height: 1.55; font-size: 14px; font-family: -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; color: inherit; }
article.markdown-body > :first-child { margin-top: 0; }
article.markdown-body > :last-child { margin-bottom: 0; }
article.markdown-body h1,article.markdown-body h2,article.markdown-body h3,article.markdown-body h4,article.markdown-body h5,article.markdown-body h6 { margin: .9em 0 .45em; font-weight: 650; color: inherit; }
article.markdown-body h1 { font-size: 1.25rem; }
article.markdown-body h2 { font-size: 1.1rem; border-bottom: 1px solid rgba(17,24,39,.14); padding-bottom: .25rem; }
article.markdown-body h3 { font-size: 1.0rem; }
article.markdown-body p { margin: .55em 0; }
article.markdown-body a { color: #2563eb; text-decoration: underline; text-underline-offset: .15em; }
article.markdown-body a:hover { text-decoration-thickness: 2px; }
article.markdown-body ul,article.markdown-body ol { margin: .4em 0; padding-left: 1.2em; }
article.markdown-body li { margin: .2em 0; }
article.markdown-body code { background: rgba(15,23,42,.08); padding: .15em .35em; border-radius: 6px; font-family: ui-monospace, SFMono-Regular, Consolas, 'Courier New', monospace; }
article.markdown-body pre { background: rgba(15,23,42,.06); border: 1px solid rgba(17,24,39,.12); padding: 10px 12px; border-radius: 10px; overflow: auto; }
article.markdown-body pre code { background: transparent; padding: 0; border-radius: 0; }
article.markdown-body blockquote { color: inherit; border-left: .25em solid rgba(37,99,235,.35); background: rgba(37,99,235,.06); padding: .1em .9em; margin: .5em 0; border-radius: 8px; }
article.markdown-body blockquote.callout { border-left-color: rgba(37,99,235,.55); }
article.markdown-body blockquote.callout.warning { border-left-color: rgba(217,119,6,.6); background: rgba(217,119,6,.08); }
article.markdown-body blockquote.callout.danger { border-left-color: rgba(220,38,38,.6); background: rgba(220,38,38,.08); }
article.markdown-body hr { border: 0; border-top: 1px solid rgba(17,24,39,.12); margin: .8em 0; }
article.markdown-body table { border-collapse: collapse; display: block; width: 100%; overflow: auto; margin: .6em 0; border: 1px solid rgba(17,24,39,.12); border-radius: 10px; }
article.markdown-body th, article.markdown-body td { border-bottom: 1px solid rgba(17,24,39,.10); padding: 6px 10px; vertical-align: top; }
article.markdown-body thead th { background: rgba(15,23,42,.06); font-weight: 650; }
article.markdown-body tbody tr:nth-child(2n) { background: rgba(15,23,42,.03); }
article.markdown-body img { max-width: 100%; height: auto; border-radius: 10px; }
article.markdown-body .task-list-item { list-style: none; }
article.markdown-body .task-list-item-checkbox { margin: 0 .4em 0 0; vertical-align: middle; }
/* Optional chat bubble wrappers (opt-in; only applies when host wraps content). */
.omd-chat-row { display: flex; gap: 10px; width: 100%; }
.omd-chat-row.omd-role-user { justify-content: flex-end; }
.omd-chat-row.omd-role-assistant { justify-content: flex-start; }
.omd-chat-row.omd-role-system { justify-content: center; }
.omd-chat-bubble { max-width: min(720px, 100%); padding: 10px 12px; border-radius: 14px; background: rgba(15,23,42,.04); border: 1px solid rgba(17,24,39,.10); }
.omd-chat-row.omd-role-user .omd-chat-bubble { background: rgba(37,99,235,.10); border-color: rgba(37,99,235,.18); }
.omd-chat-row.omd-role-system .omd-chat-bubble { background: rgba(17,24,39,.06); border-color: rgba(17,24,39,.10); }
";

    private const string ChatDarkCss = @"
/* omd-chat */
html,body { height: 100%; }
body { background: transparent; color: #e5e7eb; margin: 0; }
article.markdown-body { max-width: 100%; margin: 0; padding: 0; line-height: 1.55; font-size: 14px; font-family: -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; color: inherit; }
article.markdown-body > :first-child { margin-top: 0; }
article.markdown-body > :last-child { margin-bottom: 0; }
article.markdown-body h1,article.markdown-body h2,article.markdown-body h3,article.markdown-body h4,article.markdown-body h5,article.markdown-body h6 { margin: .9em 0 .45em; font-weight: 650; color: inherit; }
article.markdown-body h1 { font-size: 1.25rem; }
article.markdown-body h2 { font-size: 1.1rem; border-bottom: 1px solid rgba(229,231,235,.18); padding-bottom: .25rem; }
article.markdown-body h3 { font-size: 1.0rem; }
article.markdown-body p { margin: .55em 0; }
article.markdown-body a { color: #93c5fd; text-decoration: underline; text-underline-offset: .15em; }
article.markdown-body a:hover { text-decoration-thickness: 2px; }
article.markdown-body ul,article.markdown-body ol { margin: .4em 0; padding-left: 1.2em; }
article.markdown-body li { margin: .2em 0; }
article.markdown-body code { background: rgba(148,163,184,.18); padding: .15em .35em; border-radius: 6px; font-family: ui-monospace, SFMono-Regular, Consolas, 'Courier New', monospace; }
article.markdown-body pre { background: rgba(148,163,184,.10); border: 1px solid rgba(229,231,235,.16); padding: 10px 12px; border-radius: 10px; overflow: auto; }
article.markdown-body pre code { background: transparent; padding: 0; border-radius: 0; }
article.markdown-body blockquote { color: inherit; border-left: .25em solid rgba(147,197,253,.45); background: rgba(59,130,246,.12); padding: .1em .9em; margin: .5em 0; border-radius: 8px; }
article.markdown-body blockquote.callout { border-left-color: rgba(147,197,253,.6); }
article.markdown-body blockquote.callout.warning { border-left-color: rgba(251,191,36,.6); background: rgba(251,191,36,.10); }
article.markdown-body blockquote.callout.danger { border-left-color: rgba(248,113,113,.6); background: rgba(248,113,113,.10); }
article.markdown-body hr { border: 0; border-top: 1px solid rgba(229,231,235,.14); margin: .8em 0; }
article.markdown-body table { border-collapse: collapse; display: block; width: 100%; overflow: auto; margin: .6em 0; border: 1px solid rgba(229,231,235,.14); border-radius: 10px; }
article.markdown-body th, article.markdown-body td { border-bottom: 1px solid rgba(229,231,235,.12); padding: 6px 10px; vertical-align: top; }
article.markdown-body thead th { background: rgba(148,163,184,.10); font-weight: 650; }
article.markdown-body tbody tr:nth-child(2n) { background: rgba(148,163,184,.06); }
article.markdown-body img { max-width: 100%; height: auto; border-radius: 10px; }
article.markdown-body .task-list-item { list-style: none; }
article.markdown-body .task-list-item-checkbox { margin: 0 .4em 0 0; vertical-align: middle; }
/* Optional chat bubble wrappers (opt-in; only applies when host wraps content). */
.omd-chat-row { display: flex; gap: 10px; width: 100%; }
.omd-chat-row.omd-role-user { justify-content: flex-end; }
.omd-chat-row.omd-role-assistant { justify-content: flex-start; }
.omd-chat-row.omd-role-system { justify-content: center; }
.omd-chat-bubble { max-width: min(720px, 100%); padding: 10px 12px; border-radius: 14px; background: rgba(148,163,184,.10); border: 1px solid rgba(229,231,235,.14); }
.omd-chat-row.omd-role-user .omd-chat-bubble { background: rgba(59,130,246,.20); border-color: rgba(147,197,253,.22); }
.omd-chat-row.omd-role-system .omd-chat-bubble { background: rgba(148,163,184,.12); border-color: rgba(229,231,235,.14); }
";

    private const string ChatAutoCss = @"
/* omd-chat */
" + ChatLightCss + @"
@media (prefers-color-scheme: dark) {
" + ChatDarkCss + @"
}
/* data-theme overrides */
html[data-theme=dark] body { background: transparent; color: #e5e7eb; }
html[data-theme=dark] article.markdown-body { color: inherit; background: transparent; }
html[data-theme=light] body { background: transparent; color: #111827; }
html[data-theme=light] article.markdown-body { color: inherit; background: transparent; }
";

    // GitHub-ish light
    private const string GithubLightCss = @"
:root { color-scheme: light; }
html,body { height: 100%; }
body { background: #ffffff; color: #24292f; margin: 0; }
article.markdown-body { max-width: 980px; margin: 2rem auto; padding: 0 1rem; font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, ""Apple Color Emoji"", ""Segoe UI Emoji""; color: inherit; }
article.markdown-body h1,article.markdown-body h2,article.markdown-body h3 { font-weight: 600; }
article.markdown-body h1 { font-size: 2rem; }
article.markdown-body h2 { font-size: 1.5rem; border-bottom: 1px solid #d8dee4; padding-bottom: .3rem; }
article.markdown-body h3 { font-size: 1.25rem; }
article.markdown-body a { color: #0969da; text-decoration: underline; text-underline-offset: .15em; }
article.markdown-body a:hover { text-decoration-thickness: 2px; }
article.markdown-body table { width: 100%; border-collapse: collapse; }
article.markdown-body th, article.markdown-body td { border: 1px solid #d0d7de; padding: 6px 13px; }
article.markdown-body tr:nth-child(2n) { background-color: #f6f8fa; }
article.markdown-body code { background: rgba(175,184,193,.2); padding: .2em .4em; border-radius: 6px; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, ""Liberation Mono"", ""Courier New"", monospace; font-size: .9em; }
article.markdown-body pre { background: #f6f8fa; padding: 12px; border-radius: 6px; overflow: auto; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, ""Liberation Mono"", ""Courier New"", monospace; }
article.markdown-body pre code { background: none; padding: 0; border-radius: 0; font-size: inherit; }
article.markdown-body blockquote { color: #57606a; border-left: .25em solid #d0d7de; padding: 0 1em; margin: 0; }
.anchor { margin-right: .25em; opacity: .6; text-decoration: none; }
.theme-toggle { position: fixed; top: 12px; right: 12px; z-index: 9999; border: 1px solid rgba(27,31,36,.15); background: rgba(240,246,252,.9); color: inherit; border-radius: 6px; padding: 6px 8px; cursor: pointer; }
/* anchor + back-to-top UX */
.heading-anchor { margin-left: .4rem; opacity: 0; text-decoration: none; font-size: .9em; }
article.markdown-body h1:hover .heading-anchor,
article.markdown-body h2:hover .heading-anchor,
article.markdown-body h3:hover .heading-anchor,
article.markdown-body h4:hover .heading-anchor,
article.markdown-body h5:hover .heading-anchor,
article.markdown-body h6:hover .heading-anchor { opacity: .8; }
.back-to-top { margin: .25rem 0 0 0; }
.back-to-top a { font-size: .85em; opacity: .8; }
details.md-toc { margin: .5rem 0 1rem 0; }
details.md-toc > summary { cursor: pointer; font-weight: 600; margin-bottom: .5rem; }
";

    // GitHub-ish dark
    private const string GithubDarkCss = @"
:root { color-scheme: dark; }
html,body { height: 100%; }
body { background: #0d1117; color: #c9d1d9; margin: 0; }
article.markdown-body { max-width: 980px; margin: 2rem auto; padding: 0 1rem; font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, ""Apple Color Emoji"", ""Segoe UI Emoji""; color: inherit; background: transparent; }
article.markdown-body h1,article.markdown-body h2,article.markdown-body h3 { font-weight: 600; color: #e6edf3; }
article.markdown-body h2 { border-bottom: 1px solid #30363d; padding-bottom: .3rem; }
article.markdown-body a { color: #2f81f7; text-decoration: underline; text-underline-offset: .15em; }
article.markdown-body table { width: 100%; border-collapse: collapse; }
article.markdown-body th, article.markdown-body td { border: 1px solid #30363d; padding: 6px 13px; }
article.markdown-body tr:nth-child(2n) { background-color: #161b22; }
article.markdown-body code { background: #161b22; padding: .2em .4em; border-radius: 6px; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, ""Liberation Mono"", ""Courier New"", monospace; font-size: .9em; }
article.markdown-body pre { background: #161b22; padding: 12px; border-radius: 6px; overflow: auto; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, ""Liberation Mono"", ""Courier New"", monospace; }
article.markdown-body pre code { background: none; padding: 0; border-radius: 0; font-size: inherit; }
article.markdown-body blockquote { color: #8b949e; border-left: .25em solid #30363d; padding: 0 1em; margin: 0; }
.anchor { margin-right: .25em; opacity: .6; text-decoration: none; }
.theme-toggle { position: fixed; top: 12px; right: 12px; z-index: 9999; border: 1px solid rgba(240,246,252,.2); background: rgba(22,27,34,.85); color: inherit; border-radius: 6px; padding: 6px 8px; cursor: pointer; }
/* anchor + back-to-top UX */
.heading-anchor { margin-left: .4rem; opacity: 0; text-decoration: none; font-size: .9em; }
article.markdown-body h1:hover .heading-anchor,
article.markdown-body h2:hover .heading-anchor,
article.markdown-body h3:hover .heading-anchor,
article.markdown-body h4:hover .heading-anchor,
article.markdown-body h5:hover .heading-anchor,
article.markdown-body h6:hover .heading-anchor { opacity: .8; }
.back-to-top { margin: .25rem 0 0 0; }
.back-to-top a { font-size: .85em; opacity: .85; }
details.md-toc { margin: 0 0 1rem 0; }
details.md-toc > summary { cursor: pointer; font-weight: 600; margin-bottom: .5rem; }
";

    // Auto via prefers-color-scheme
    private const string GithubAutoCss = @"
/* light defaults */
html,body { height: 100%; }
body { background: #ffffff; color: #24292f; margin: 0; }
article.markdown-body { max-width: 980px; margin: 2rem auto; padding: 0 1rem; font-family: ui-sans-serif, system-ui, -apple-system, Segoe UI, Roboto, Helvetica, Arial, ""Apple Color Emoji"", ""Segoe UI Emoji""; color: inherit; }
article.markdown-body h1,article.markdown-body h2,article.markdown-body h3 { font-weight: 600; }
article.markdown-body h2 { border-bottom: 1px solid #d8dee4; padding-bottom: .3rem; }
article.markdown-body a { color: #0969da; text-decoration: underline; text-underline-offset: .15em; }
article.markdown-body table { width: 100%; border-collapse: collapse; }
article.markdown-body th, article.markdown-body td { border: 1px solid #d0d7de; padding: 6px 13px; }
article.markdown-body tr:nth-child(2n) { background-color: #f6f8fa; }
article.markdown-body code { background: rgba(175,184,193,.2); padding: .2em .4em; border-radius: 6px; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, ""Liberation Mono"", ""Courier New"", monospace; font-size: .9em; }
article.markdown-body pre { background: #f6f8fa; padding: 12px; border-radius: 6px; overflow: auto; font-family: ui-monospace, SFMono-Regular, Menlo, Consolas, ""Liberation Mono"", ""Courier New"", monospace; }
article.markdown-body pre code { background: none; padding: 0; border-radius: 0; font-size: inherit; }
article.markdown-body blockquote { color: #57606a; border-left: .25em solid #d0d7de; padding: 0 1em; margin: 0; }
.anchor { margin-right: .25em; opacity: .6; text-decoration: none; }

@media (prefers-color-scheme: dark) {
  body { background: #0d1117; color: #c9d1d9; }
  article.markdown-body { color: inherit; background: transparent; }
  article.markdown-body h1,article.markdown-body h2,article.markdown-body h3 { color: #e6edf3; }
  article.markdown-body h2 { border-bottom: 1px solid #30363d; }
  article.markdown-body a { color: #2f81f7; }
  article.markdown-body th, article.markdown-body td { border-color: #30363d; }
  article.markdown-body tr:nth-child(2n) { background-color: #161b22; }
  article.markdown-body code, article.markdown-body pre { background: #161b22; }
  article.markdown-body pre code { background: none; }
  article.markdown-body blockquote { color: #8b949e; border-left-color: #30363d; }
}
/* data-theme overrides */
html[data-theme=dark] body { background: #0d1117; color: #c9d1d9; }
html[data-theme=dark] article.markdown-body { color: inherit; background: transparent; }
html[data-theme=light] body { background: #ffffff; color: #24292f; }
.theme-toggle { position: fixed; top: 12px; right: 12px; z-index: 9999; border: 1px solid rgba(27,31,36,.15); background: rgba(240,246,252,.9); color: inherit; border-radius: 6px; padding: 6px 8px; cursor: pointer; }
@media (prefers-color-scheme: dark) { .theme-toggle { border-color: rgba(240,246,252,.2); background: rgba(22,27,34,.85); } }
html[data-theme=dark] .theme-toggle { border-color: rgba(240,246,252,.2); background: rgba(22,27,34,.85); }
/* anchor + back-to-top UX */
.heading-anchor { margin-left: .4rem; opacity: 0; text-decoration: none; font-size: .9em; }
article.markdown-body h1:hover .heading-anchor,
article.markdown-body h2:hover .heading-anchor,
article.markdown-body h3:hover .heading-anchor,
article.markdown-body h4:hover .heading-anchor,
article.markdown-body h5:hover .heading-anchor,
article.markdown-body h6:hover .heading-anchor { opacity: .8; }
.back-to-top { margin: .25rem 0 0 0; }
.back-to-top a { font-size: .85em; opacity: .8; }
details.md-toc { margin: 0 0 1rem 0; }
details.md-toc > summary { cursor: pointer; font-weight: 600; margin-bottom: .5rem; }
";

    internal static string ThemeToggleScript => @"
(function(){
  var btn = document.querySelector('[data-theme-toggle]');
  if(!btn) return;
  function set(t){ document.documentElement.setAttribute('data-theme', t); try{ localStorage.setItem('md-theme', t);}catch(e){} }
  try{ var saved = localStorage.getItem('md-theme'); if(saved){ set(saved); } }catch(e){}
  btn.addEventListener('click', function(){ var cur = document.documentElement.getAttribute('data-theme')||'auto'; set(cur==='dark'?'light':'dark'); });
})();
";

    internal static string AnchorCopyScript => @"
(function(){
  function copy(text){ try{ navigator.clipboard.writeText(text); }catch(e){ var ta=document.createElement('textarea'); ta.value=text; document.body.appendChild(ta); ta.select(); try{ document.execCommand('copy'); }finally{ document.body.removeChild(ta);} } }
  function buildUrl(id){ try{ var u=new URL(window.location.href); u.hash = id ? ('#'+id) : ''; return u.toString(); }catch(e){ return '#'+id; } }
  document.addEventListener('click', function(ev){
    var a = ev.target.closest && ev.target.closest('a.heading-anchor');
    if(!a) return; ev.preventDefault();
    var id = a.getAttribute('data-anchor-id');
    if(!id) return; copy(buildUrl(id));
    a.setAttribute('data-copied','true'); setTimeout(function(){ a.removeAttribute('data-copied'); }, 1200);
  }, false);
})();
";

    // Additional CSS shared across styles for enhanced TOC/ScrollSpy. Scoped by ScopeCss() via explicit selectors.
    internal const string CommonExtraCss = @"
/* Printed pages: headings and lead-in paragraphs stay with what follows, and short blocks are not split */
@media print {
  article.markdown-body h1, article.markdown-body h2, article.markdown-body h3, article.markdown-body h4, article.markdown-body h5, article.markdown-body h6 { break-after: avoid; }
  article.markdown-body p + pre, article.markdown-body p + ul, article.markdown-body p + ol, article.markdown-body p + table, article.markdown-body p + figure, article.markdown-body li > pre { break-before: avoid; }
  article.markdown-body pre, article.markdown-body tr, article.markdown-body figure, article.markdown-body blockquote, article.markdown-body img { break-inside: avoid; }
}
/* Enhanced TOC */
article.markdown-body nav.md-toc { font-size: .95rem; line-height: 1.45; background: #f6f8fa; border: 1px solid #d0d7de; border-radius: 8px; padding: 10px 12px; margin: .5rem 0 1rem 0; }
article.markdown-body nav.md-toc .toc-title { font-weight: 600; margin: 0 0 .25rem 0; font-size: 1rem; }
article.markdown-body nav.md-toc ul { list-style: none; margin: 0; padding-left: 0; }
article.markdown-body nav.md-toc > ul { margin-left: 0; padding-left: 0; }
article.markdown-body nav.md-toc > ul > li { margin-left: 0; }
article.markdown-body nav.md-toc li { margin: .15rem 0; }
/* Make nesting visually obvious */
article.markdown-body nav.md-toc ul ul { list-style: disc; margin-left: 1rem; padding-left: 1rem; border-left: none; }
article.markdown-body nav.md-toc ul ul li { margin-left: .1rem; }
article.markdown-body nav.md-toc a { color: inherit; text-decoration: none; }
article.markdown-body nav.md-toc a:hover { text-decoration: underline; }
article.markdown-body nav.md-toc a.active { color: #0969da; font-weight: 600; border-left: 2px solid #0969da; padding-left: .4rem; }
/* Panel variant */
article.markdown-body nav.md-toc.panel { box-shadow: 0 1px 0 rgba(27,31,36,.04); }
article.markdown-body nav.md-toc.no-chrome { background: transparent; border: none; box-shadow: none; }
article.markdown-body nav.md-toc.outline { background: transparent; }
/* no extra borders for nested items in simplified chrome */
/* Sidebar variants */
article.markdown-body nav.md-toc.sidebar { max-height: calc(100vh - 2rem); overflow: auto; }
article.markdown-body nav.md-toc.sidebar.sticky { position: sticky; top: 1rem; }
article.markdown-body nav.md-toc.sidebar.right { float: right; width: 260px; margin: 0 0 1rem 1rem; }
article.markdown-body nav.md-toc.sidebar.left  { float: left;  width: 260px; margin: 0 1rem 1rem 0; }
/* Two-column grid layout overrides floats */
article.markdown-body .md-layout.two-col { display: grid; grid-template-columns: var(--md-toc-width, 260px) 1fr; gap: 1rem; align-items: start; }
article.markdown-body .md-layout.two-col.right { grid-template-columns: 1fr var(--md-toc-width, 260px); }
article.markdown-body .md-layout.two-col > nav.md-toc.sidebar { float: none; width: auto; margin: 0; }
article.markdown-body .md-layout.two-col .md-content { min-width: 0; }
article.markdown-body .md-layout.two-col > nav.md-toc.sidebar.sticky { position: sticky; top: 1rem; max-height: calc(100vh - 2rem); overflow: auto; }
/* Images never overflow the text column */
article.markdown-body img { max-width: 100%; }
/* Comfortable vertical rhythm for block elements */
article.markdown-body table { margin: .75rem 0 1.1rem 0; }
article.markdown-body pre { margin: .75rem 0 1.1rem 0; }
article.markdown-body blockquote { margin: .65rem 0 1rem 0; }

/* Task lists (GitHub-like) */
/* Only task items drop their marker; ordinary items in the same list keep their bullets. */
article.markdown-body li.task-list-item { list-style: none; }
article.markdown-body input.task-list-item-checkbox { margin: 0 .35em .15em -1.4em; vertical-align: middle; }
@media (max-width: 1000px) {
  article.markdown-body nav.md-toc.sidebar.right,
  article.markdown-body nav.md-toc.sidebar.left { float: none; width: auto; margin: 0 0 1rem 0; }
  article.markdown-body .md-layout.two-col { display: block; }
  article.markdown-body nav.md-toc.hide-narrow { display: none; }
}
/* Dark-mode adjustments (prefers + data-theme) */
@media (prefers-color-scheme: dark) {
  article.markdown-body nav.md-toc { background: #161b22; border-color: #30363d; }
  article.markdown-body nav.md-toc ul ul { border-left-color: #30363d; }
  article.markdown-body nav.md-toc a.active { color: #2f81f7; border-left-color: #2f81f7; }
}
html[data-theme=dark] article.markdown-body nav.md-toc { background: #161b22; border-color: #30363d; }
html[data-theme=dark] article.markdown-body nav.md-toc ul ul { border-left-color: #30363d; }
html[data-theme=dark] article.markdown-body nav.md-toc a.active { color: #2f81f7; border-left-color: #2f81f7; }
/* Explicit light-mode overrides to beat prefers-color-scheme when user forces light */
html[data-theme=light] article.markdown-body nav.md-toc { background: #f6f8fa; border-color: #d0d7de; }
html[data-theme=light] article.markdown-body nav.md-toc ul ul { border-left-color: rgba(27,31,36,.12); }
html[data-theme=light] article.markdown-body nav.md-toc a { color: inherit; }
html[data-theme=light] article.markdown-body h1,
html[data-theme=light] article.markdown-body h2,
html[data-theme=light] article.markdown-body h3,
html[data-theme=light] article.markdown-body h4,
html[data-theme=light] article.markdown-body h5,
html[data-theme=light] article.markdown-body h6 { color: #24292f; }
/* Force light zebra, borders, and code/pre backgrounds in light mode */
html[data-theme=light] article.markdown-body tr:nth-child(2n) { background-color: #f6f8fa; }
html[data-theme=light] article.markdown-body th, html[data-theme=light] article.markdown-body td { border-color: #d0d7de; }
html[data-theme=light] article.markdown-body pre { background: #f6f8fa; }
html[data-theme=light] article.markdown-body code { background: rgba(175,184,193,.2); }
html[data-theme=light] article.markdown-body pre code { background: none; }
/* Robust wrapping like GitHub: break long unbroken strings only when they cannot fit.
   overflow-wrap: break-word, unlike anywhere, does not shrink table columns below their longest word. */
article.markdown-body,
article.markdown-body .md-content { overflow-wrap: break-word; word-wrap: break-word; }
article.markdown-body h1,
article.markdown-body h2,
article.markdown-body h3,
article.markdown-body h4,
article.markdown-body h5,
article.markdown-body h6,
article.markdown-body p,
article.markdown-body li,
article.markdown-body blockquote,
article.markdown-body td,
article.markdown-body th { overflow-wrap: break-word; }

/* Footnotes */
article.markdown-body section.footnotes { margin-top: 2rem; }
article.markdown-body section.footnotes > hr { border: 0; border-top: 1px solid rgba(27,31,36,.15); margin: 1rem 0; }
article.markdown-body section.footnotes ol { margin: 0; padding-left: 1.25rem; }
article.markdown-body section.footnotes li { margin: .35rem 0; }
article.markdown-body a.footnote-backref { margin-left: .5rem; text-decoration: none; opacity: .8; }
article.markdown-body a.footnote-backref:hover { opacity: 1; text-decoration: underline; }
@media (prefers-color-scheme: dark) {
  article.markdown-body section.footnotes > hr { border-top-color: rgba(240,246,252,.2); }
}
html[data-theme=dark] article.markdown-body section.footnotes > hr { border-top-color: rgba(240,246,252,.2); }

/* Escaped raw HTML blocks */
article.markdown-body pre.md-raw-html { white-space: pre-wrap; }

/* Copy buttons (optional; used by OfficeIMO.MarkdownRenderer shell helpers) */
article.markdown-body pre.omd-has-actions { position: relative; }
article.markdown-body .omd-copy-btn {
  font: inherit;
  font-size: 12px;
  line-height: 1;
  padding: 6px 8px;
  border-radius: 8px;
  border: 1px solid rgba(27,31,36,.15);
  background: rgba(255,255,255,.78);
  color: inherit;
  cursor: pointer;
}
article.markdown-body .omd-copy-btn:hover { background: rgba(255,255,255,.92); }
article.markdown-body .omd-copy-btn[data-omd-flash] { border-color: rgba(9,105,218,.35); }
article.markdown-body pre.omd-has-actions > .omd-copy-code { position: absolute; top: 10px; right: 10px; z-index: 2; }
article.markdown-body .omd-table-actions { display: flex; gap: 8px; justify-content: flex-end; margin: .15rem 0 .35rem 0; }

@media (prefers-color-scheme: dark) {
  article.markdown-body .omd-copy-btn { border-color: rgba(240,246,252,.2); background: rgba(22,27,34,.78); }
  article.markdown-body .omd-copy-btn:hover { background: rgba(22,27,34,.92); }
}
html[data-theme=dark] article.markdown-body .omd-copy-btn { border-color: rgba(240,246,252,.2); background: rgba(22,27,34,.78); }
html[data-theme=dark] article.markdown-body .omd-copy-btn:hover { background: rgba(22,27,34,.92); }
";

    // Word-like, document-centric styling – Calibri/Cambria fonts, comfortable spacing,
    // and Word-ish table formatting (header shading, banded rows, clear borders).
    private const string WordCss = @"
/* Base */
html,body { height:100%; }
body { background:#ffffff; color:#1f1f1f; margin:0; }
article.markdown-body { max-width: 900px; margin: 1.5rem auto; padding: 0 1rem; line-height: 1.5; font-family: Calibri, 'Segoe UI', Roboto, Helvetica, Arial, sans-serif; }
article.markdown-body h1, article.markdown-body h2, article.markdown-body h3,
article.markdown-body h4, article.markdown-body h5, article.markdown-body h6 {
  font-family: Cambria, 'Times New Roman', Times, serif; color: #2F5496; margin: 1.2em 0 .5em;
}
article.markdown-body h1 { font-size: 2.0rem; font-weight: 600; color:#1f3763; }
article.markdown-body h2 { font-size: 1.6rem; font-weight: 600; border-bottom: 1px solid #d6d6d6; padding-bottom: .25rem; }
article.markdown-body h3 { font-size: 1.3rem; font-weight: 600; }
article.markdown-body p { margin: .6em 0; }
article.markdown-body a { color: #2b579a; text-decoration: underline; text-underline-offset: .12em; }
article.markdown-body a:hover { text-decoration-thickness: 2px; }

/* Lists – slightly larger indent, consistent spacing */
article.markdown-body ul, article.markdown-body ol { margin: .4rem 0 .8rem 1.6rem; }
article.markdown-body li { margin: .2rem 0; }

/* Code */
article.markdown-body code { background: #f2f2f2; padding: .15em .35em; border-radius: 4px; font-family: Consolas, Menlo, 'Liberation Mono', 'Courier New', monospace; }
article.markdown-body pre { background: #f2f2f2; padding: 12px; border-radius: 6px; overflow: auto; font-family: Consolas, Menlo, 'Liberation Mono', 'Courier New', monospace; }
article.markdown-body pre code { background: none; padding: 0; border-radius: 0; }

/* Tables – mimic Word's Grid Table style */
article.markdown-body table { width: 100%; border-collapse: collapse; border: 1px solid #d6d6d6; margin: .8rem 0 1.2rem; }
article.markdown-body thead th { background: #e7e6e6; color: #1f1f1f; font-weight: 600; }
article.markdown-body th, article.markdown-body td { border: 1px solid #d6d6d6; padding: 6px 10px; vertical-align: top; }
article.markdown-body tbody tr:nth-child(2n) { background: #f8f8f8; }

article.markdown-body img { max-width: 100%; height: auto; }

/* Blockquote – subtle Word-like bar */
article.markdown-body blockquote { color: #404040; border-left: .25em solid #c8c8c8; padding: 0 1em; margin: .4rem 0 .9rem; }

/* Anchors + back-to-top */
.heading-anchor { margin-left: .35rem; opacity: .65; text-decoration: none; font-size: .9em; color: #2b579a; }
.back-to-top { margin: .25rem 0 0; }
.back-to-top a { font-size: .85em; opacity: .85; color: #2b579a; }

/* Print – repeat table header row like Word */
@media print {
  article.markdown-body thead { display: table-header-group; }
  article.markdown-body tr { break-inside: avoid; }
}

/* Explicit wrapping for long tokens */
article.markdown-body, article.markdown-body .md-content { overflow-wrap:break-word; }
article.markdown-body td, article.markdown-body th, article.markdown-body p, article.markdown-body li { overflow-wrap:break-word; }
";

    internal static string ScrollSpyScript => @"
(function(){
  var scope = document.querySelector('article.markdown-body') || document;
  var navs = scope.querySelectorAll('nav.md-toc[data-md-scrollspy=""1""], nav.md-toc.md-scrollspy');
  if(!navs || navs.length===0) return;
  var links = [];
  navs.forEach(function(nav){ links = links.concat(Array.from(nav.querySelectorAll('a[href^=""#""]'))); });
  if(links.length===0) return;
  var map = new Map();
  var heads = [];
  links.forEach(function(a){ try{ var id = decodeURIComponent(a.getAttribute('href').slice(1)); var el = document.getElementById(id); if(el){ map.set(el.id, a); heads.push(el); } }catch(e){} });
  if(heads.length===0) return;
  var activeLink = null;
  var opts = { root: null, rootMargin: '0px 0px -60% 0px', threshold: 0.01 };
  var obs = new IntersectionObserver(function(entries){
    entries.forEach(function(entry){
      var id = entry.target.id; var link = map.get(id);
      if(!link) return;
      if(entry.isIntersecting){ if(activeLink){ activeLink.classList.remove('active'); }
        link.classList.add('active'); activeLink = link;
        var nav = link.closest('nav.md-toc'); if(nav && (nav.dataset.autoscroll==='1' || nav.classList.contains('autoscroll'))){
          try{ link.scrollIntoView({ block:'nearest', inline:'nearest' }); }catch(e){}
        }
      }
    });
  }, opts);
  heads.forEach(function(h){ obs.observe(h); });
})();
";
}
