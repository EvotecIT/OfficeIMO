# OfficeIMO.MarkdownRenderer.SamplePlugin

`OfficeIMO.MarkdownRenderer.SamplePlugin` is a small sample package that demonstrates how a third-party plugin can build on top of `OfficeIMO.MarkdownRenderer` and `OfficeIMO.Markdown.Html`.

It shows one complete contract:

- a fenced-block renderer plugin
- a reusable feature pack built from that plugin
- source-side reader document transforms for plugin-owned inline/AST upgrades
- HTML-to-markdown document transforms for non-shared visual HTML paths
- custom HTML element block converters for vendor HTML that never used the shared visual contract
- custom HTML inline element converters for vendor inline HTML that should recover semantic inline AST
- shared `data-omd-*` visual host HTML
- plugin-owned HTML round-trip hints for `HtmlToMarkdownOptions`

## Example

```csharp
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.SamplePlugin;

var renderOptions = MarkdownRendererPresets.CreateStrictMinimal();
renderOptions.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);

var html = MarkdownRenderer.RenderBodyHtml("""
```status-panel
{"title":"Operations Overview","summary":"All checks passing","status":"healthy","caption":"Panel caption"}
```
""", renderOptions);

var htmlToMarkdown = new HtmlToMarkdownOptions();
htmlToMarkdown.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);
var document = html.LoadFromHtml(htmlToMarkdown);

var readerOptions = MarkdownReaderOptions.CreatePortableProfile();
readerOptions.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);
var parsed = MarkdownReader.Parse("System {{status:Healthy}} now", readerOptions);
```

The sample pack also carries HTML-ingestion helpers, so both plain `<pre><code class="language-status-panel">...</code></pre>` input and vendor-specific `<section class="sample-status-panel" ...>` HTML can recover back into a semantic `status-panel` fenced block without touching core converter logic.
It also shows inline recovery through `<span class="sample-status-badge">Healthy</span>`, which round-trips into typed markdown highlight inline AST instead of raw HTML.
On the source side, it also upgrades plain `{{status:Healthy}}` tokens into typed highlight inline AST through the same plugin/feature-pack contract.

Use this package as a reference implementation when building external plugin packs that want renderer behavior, HTML ingestion, and HTML round-trip fidelity to stay aligned through the same reusable feature-pack contract.
