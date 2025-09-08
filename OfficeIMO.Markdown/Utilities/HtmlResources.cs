using System.Text;

namespace OfficeIMO.Markdown.Utilities;

/// <summary>
/// Built-in CSS and tiny scripts for HTML rendering.
/// </summary>
internal static class HtmlResources {
    internal static string GetStyleCss(HtmlStyle style) => style switch {
        HtmlStyle.Plain => string.Empty,
        HtmlStyle.Clean => CleanCss,
        HtmlStyle.GithubLight => GithubLightCss,
        HtmlStyle.GithubDark => GithubDarkCss,
        HtmlStyle.GithubAuto => GithubAutoCss,
        _ => CleanCss
    };

    // Minimal, readable defaults
    private const string CleanCss = @"
html,body { height: 100%; }
body { background: #ffffff; color: #24292f; margin: 0; }
article.markdown-body { max-width: 860px; margin: 2rem auto; padding: 0 1rem; line-height: 1.6; font-family: -apple-system, Segoe UI, Roboto, Helvetica, Arial, sans-serif; color: inherit; }
article.markdown-body h1,article.markdown-body h2,article.markdown-body h3,article.markdown-body h4 { margin-top: 1.6em; margin-bottom: .6em; font-weight: 600; }
article.markdown-body h1 { font-size: 2.0rem; }
article.markdown-body h2 { font-size: 1.6rem; border-bottom: 1px solid #eaecef; padding-bottom: .3rem; }
article.markdown-body h3 { font-size: 1.25rem; }
article.markdown-body p { margin: .8em 0; }
article.markdown-body a { color: #0969da; text-decoration: underline; text-underline-offset: .15em; }
article.markdown-body a:hover { text-decoration-thickness: 2px; }
article.markdown-body table { border-collapse: collapse; display: block; width: 100%; overflow: auto; }
article.markdown-body th, article.markdown-body td { border: 1px solid #d0d7de; padding: 6px 13px; }
article.markdown-body tr:nth-child(2n) { background-color: #f6f8fa; }
article.markdown-body code { background: rgba(175,184,193,.2); padding: .2em .4em; border-radius: 6px; }
article.markdown-body pre { background: #f6f8fa; padding: 12px; border-radius: 6px; overflow: auto; }
article.markdown-body blockquote { color: #57606a; border-left: .25em solid #d0d7de; padding: 0 1em; margin: 0; }
article.markdown-body blockquote.callout { border-left-color: #0969da; background: #f6f8fa; }
article.markdown-body blockquote.callout.warning { border-left-color: #d4a72c; }
article.markdown-body blockquote.callout.danger { border-left-color: #cf222e; }
.anchor { margin-right: .25em; opacity: .6; text-decoration: none; }
/* data-theme overrides */
html[data-theme=dark] body { background: #0d1117; color: #c9d1d9; }
html[data-theme=light] body { background: #ffffff; color: #24292f; }
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
article.markdown-body code { background: rgba(175,184,193,.2); padding: .2em .4em; border-radius: 6px; }
article.markdown-body pre { background: #f6f8fa; padding: 12px; border-radius: 6px; overflow: auto; }
article.markdown-body blockquote { color: #57606a; border-left: .25em solid #d0d7de; padding: 0 1em; margin: 0; }
.anchor { margin-right: .25em; opacity: .6; text-decoration: none; }
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
article.markdown-body code { background: #161b22; padding: .2em .4em; border-radius: 6px; }
article.markdown-body pre { background: #161b22; padding: 12px; border-radius: 6px; overflow: auto; }
article.markdown-body blockquote { color: #8b949e; border-left: .25em solid #30363d; padding: 0 1em; margin: 0; }
.anchor { margin-right: .25em; opacity: .6; text-decoration: none; }
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
article.markdown-body code { background: rgba(175,184,193,.2); padding: .2em .4em; border-radius: 6px; }
article.markdown-body pre { background: #f6f8fa; padding: 12px; border-radius: 6px; overflow: auto; }
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
  article.markdown-body blockquote { color: #8b949e; border-left-color: #30363d; }
}
/* data-theme overrides */
html[data-theme=dark] body { background: #0d1117; color: #c9d1d9; }
html[data-theme=dark] article.markdown-body { color: inherit; background: transparent; }
html[data-theme=light] body { background: #ffffff; color: #24292f; }
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
}
