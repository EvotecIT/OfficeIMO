OfficeIMO.MarkdownRenderer
=========================

Small helper library to render Markdown using `OfficeIMO.Markdown` into HTML that is easy to host in WebView2 (or any browser):

- `BuildShellHtml(...)`: returns a full HTML page that preloads CSS/Prism/Mermaid once
- `RenderBodyHtml(...)`: returns an HTML fragment for a given Markdown string
- `BuildUpdateScript(...)`: returns a JavaScript snippet calling `updateContent(...)`

Typical usage (WebView2)

```csharp
using OfficeIMO.MarkdownRenderer;

// 1) Load shell once
var opts = new MarkdownRendererOptions();
webView.NavigateToString(MarkdownRenderer.BuildShellHtml("Chat", opts));

// 2) For each message update
string bodyHtml = MarkdownRenderer.RenderBodyHtml(markdownText, opts);
await webView.ExecuteScriptAsync(MarkdownRenderer.BuildUpdateScript(bodyHtml));
```

Mermaid diagrams

Write Mermaid in fenced code blocks:

```markdown
```mermaid
flowchart LR
  A --> B
```
```

Security note

Defaults are biased for untrusted chat output:
- raw HTML parsing is disabled
- `javascript:` / `vbscript:` URLs are blocked by the reader
- `file:` URLs are blocked by default in `MarkdownRendererOptions.ReaderOptions`

