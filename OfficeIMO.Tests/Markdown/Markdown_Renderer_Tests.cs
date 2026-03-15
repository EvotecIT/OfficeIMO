using System.Globalization;
using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Renderer_Tests {
    [Fact]
    public void Reader_Disallows_Javascript_Links_ByDefault() {
        var md = "[x](javascript:alert(1))";
        var doc = MarkdownReader.Parse(md);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("javascript:", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal); // rendered as plain text
    }

    [Fact]
    public void Nested_Parsing_Respects_DisallowFileUrls() {
        var options = new MarkdownReaderOptions { DisallowFileUrls = true, HtmlBlocks = false, InlineHtml = false };
        string md = """
- outer
  > [x](file:///c:/test)
""";
        var doc = MarkdownReader.Parse(md, options);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.DoesNotContain("file:", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Shell_Contains_UpdateContent_And_Mermaid_Bootstrap() {
        var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat");
        Assert.Contains("async function updateContent", shell, StringComparison.Ordinal);
        Assert.Contains("omdRoot", shell, StringComparison.Ordinal);
        Assert.Contains("mermaid.esm.min.mjs", shell, StringComparison.Ordinal);
        Assert.Contains("katex", shell, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("renderMathInElement", shell, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownRenderer_Shell_Contains_WebView2_Message_Listener() {
        var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat");
        Assert.Contains("chrome.webview", shell, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("addEventListener('message'", shell, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownRenderer_Shell_Includes_Csp_Meta_When_Configured() {
        var opts = new MarkdownRendererOptions {
            ContentSecurityPolicy = "default-src 'self'; img-src https:; style-src 'unsafe-inline' https:; script-src 'unsafe-inline' https:"
        };
        var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat", opts);
        Assert.Contains("http-equiv=\"Content-Security-Policy\"", shell, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("default-src", shell, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownRenderer_Shell_Appends_Custom_ShellCss() {
        var opts = new MarkdownRendererOptions { ShellCss = "#omdRoot { padding: 12px; }" };
        var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat", opts);
        Assert.Contains("data-omd=\"shell\"", shell, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("#omdRoot { padding: 12px; }", shell, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Shell_Contains_CopyButton_Scripts_When_Enabled() {
        var opts = new MarkdownRendererOptions { EnableCodeCopyButtons = true, EnableTableCopyButtons = true };
        var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat", opts);
        Assert.Contains("omdSetupCodeCopyButtons", shell, StringComparison.Ordinal);
        Assert.Contains("omdSetupTableCopyButtons", shell, StringComparison.Ordinal);
        Assert.Contains("omdCopyText", shell, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Can_Truncate_Markdown_When_Limit_Exceeded() {
        var opts = new MarkdownRendererOptions {
            MaxMarkdownChars = 5,
            MarkdownOverflowHandling = OverflowHandling.Truncate,
            HtmlOptions = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null }
        };
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("0123456789", opts);
        Assert.Contains("01234", html, StringComparison.Ordinal);
        Assert.DoesNotContain("56789", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Renders_Overflow_Warning_When_Html_Too_Large() {
        var opts = new MarkdownRendererOptions {
            MaxBodyHtmlBytes = 1,
            BodyHtmlOverflowHandling = OverflowHandling.RenderError,
            HtmlOptions = new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = "markdown-body" }
        };
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("Hello", opts);
        Assert.Contains("data-omd=\"overflow\"", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownRenderer_Adds_Mermaid_Hash_Attributes() {
        var md = "```mermaid\nflowchart LR\nA-->B\n```";
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md);
        Assert.Contains("class=\"mermaid\"", html, StringComparison.Ordinal);
        Assert.Contains("data-mermaid-hash", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_StrictPreset_Preserves_Safe_Underline_Tags() {
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(
            "<u>hello</u>",
            MarkdownRendererPresets.CreateStrict());

        Assert.Contains("<u>hello</u>", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("&lt;u&gt;hello&lt;/u&gt;", html, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownRenderer_StrictPreset_Still_Encodes_Unsupported_Inline_Html() {
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(
            "<span>hello</span>",
            MarkdownRendererPresets.CreateStrict());

        Assert.Contains("&lt;span&gt;hello&lt;/span&gt;", html, StringComparison.OrdinalIgnoreCase);
        Assert.DoesNotContain("<span>hello</span>", html, StringComparison.OrdinalIgnoreCase);
    }

    [Theory]
    [InlineData("chart")]
    public void MarkdownRenderer_Converts_Chart_Code_Fences_When_Enabled(string language) {
        var configJson = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var md = $"```{language}\n{configJson}\n```";
        var opts = new MarkdownRendererOptions();
        opts.Chart.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);
        Assert.Contains("canvas", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("class=\"omd-visual omd-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-contract=\"v1\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-kind=\"chart\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-omd-fence-language=\"{language}\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-hash", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-config-format=\"json\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-config-encoding=\"base64-utf8\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-config-b64", html, StringComparison.Ordinal);
        Assert.Contains("data-chart-config-b64", html, StringComparison.Ordinal);
        Assert.Equal(configJson, DecodeBase64Attribute(html, "data-omd-config-b64"));
    }

    [Fact]
    public void MarkdownRenderer_ChatPreset_Converts_IxChart_Code_Fences_When_Enabled() {
        var configJson = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var md = $"```ix-chart\n{configJson}\n```";
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        opts.Chart.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("class=\"omd-visual omd-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"ix-chart\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_ChatPreset_Converts_IxChart_Code_Fences_Inside_List_Items_When_Enabled() {
        var configJson = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var md = $"""
- item

  ```ix-chart
  {configJson}
  ```
""";
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        opts.Chart.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("<ul>", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"ix-chart\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("language-ix-chart", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_ChatPreset_Converts_Quoted_IxChart_Code_Fences_When_Enabled() {
        var configJson = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var md = $"""
> ```ix-chart
> {configJson}
> ```
""";
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        opts.Chart.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("<blockquote>", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"ix-chart\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("language-ix-chart", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_ChatPreset_Converts_Nested_Quoted_IxChart_Code_Fences_When_Enabled() {
        var configJson = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var md = $"""
> outer
>
> > ```ix-chart
> > {configJson}
> > ```
""";
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        opts.Chart.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("<blockquote>", html, StringComparison.Ordinal);
        Assert.Contains("outer", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"ix-chart\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("language-ix-chart", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_ChatPreset_Converts_List_Quoted_IxChart_Code_Fences_When_Enabled() {
        var configJson = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var md = $"""
- item

  > ```ix-chart
  > {configJson}
  > ```
""";
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        opts.Chart.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("<ul>", html, StringComparison.Ordinal);
        Assert.Contains("<blockquote>", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"ix-chart\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("language-ix-chart", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Converts_Compact_Quoted_Mermaid_Fences_When_Normalization_Is_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };
        opts.Mermaid.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("> ```mermaidflowchart LR A-->B\n> ```", opts);

        Assert.Contains("<blockquote>", html, StringComparison.Ordinal);
        Assert.Contains("class=\"mermaid\"", html, StringComparison.Ordinal);
        Assert.Contains("flowchart LR", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Converts_Quoted_Math_Fences_When_Enabled() {
        var opts = new MarkdownRendererOptions();
        opts.Math.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("> ```math\n> x^2 + 1\n> ```", opts);

        Assert.Contains("<blockquote>", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-math\"", html, StringComparison.Ordinal);
        Assert.Contains("x^2 + 1", html, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("network")]
    [InlineData("visnetwork")]
    public void MarkdownRenderer_Converts_Network_Code_Fences_When_Enabled(string language) {
        var md = $"```{language}\n{{\"nodes\":[{{\"id\":\"A\",\"label\":\"User\"}},{{\"id\":\"B\",\"label\":\"Group\"}}],\"edges\":[{{\"from\":\"A\",\"to\":\"B\",\"label\":\"memberOf\"}}]}}\n```";
        var opts = new MarkdownRendererOptions();
        opts.Network.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);
        Assert.Contains("class=\"omd-visual omd-network\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-contract=\"v1\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-kind=\"network\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-omd-fence-language=\"{language}\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-hash", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-config-format=\"json\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-config-encoding=\"base64-utf8\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-config-b64", html, StringComparison.Ordinal);
        Assert.Contains("data-network-config-b64", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_ChatPreset_Converts_IxNetwork_Code_Fences_When_Enabled() {
        var md = """
```ix-network
{"nodes":[{"id":"A","label":"User"},{"id":"B","label":"Group"}],"edges":[{"from":"A","to":"B","label":"memberOf"}]}
```
""";
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        opts.Network.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("class=\"omd-visual omd-network\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"ix-network\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Chart_Fallback_Uses_Shared_Native_Visual_Metadata() {
        var configJson = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var opts = new MarkdownRendererOptions();
        opts.Chart.Enabled = true;
        opts.FencedCodeBlockRenderers.Clear();

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml($"```chart\n{configJson}\n```", opts);

        Assert.Contains("class=\"omd-visual omd-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-contract=\"v1\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-kind=\"chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-chart-config-b64", html, StringComparison.Ordinal);
        Assert.Equal(configJson, DecodeBase64Attribute(html, "data-omd-config-b64"));
    }

    [Fact]
    public void MarkdownVisualContract_Can_Build_Shared_Visual_Metadata_For_Hosts() {
        var raw = "{\"type\":\"bar\"}";
        var payload = MarkdownVisualContract.CreatePayload(raw);
        var html = MarkdownVisualContract.BuildElementHtml(
            "div",
            "omd-visual omd-custom",
            "custom-chart",
            "vendor-chart",
            payload,
            new KeyValuePair<string, string?>("data-custom-hash", payload.Hash));

        Assert.Contains("class=\"omd-visual omd-custom\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-omd-visual-contract=\"{MarkdownVisualContract.ContractVersion}\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-kind=\"custom-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"vendor-chart\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-omd-visual-hash=\"{MarkdownVisualContract.ComputePayloadHash(raw)}\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-custom-hash=\"{payload.Hash}\"", html, StringComparison.Ordinal);
        Assert.Equal(raw, DecodeBase64Attribute(html, "data-omd-config-b64"));
    }

    [Fact]
    public void MarkdownVisualContract_Uses_Stable_Hash_For_Equivalent_Json_Payloads() {
        var minified = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var formatted = """
{
  "data": {
    "datasets": [
      {
        "data": [ 1 ],
        "label": "Count"
      }
    ],
    "labels": [ "A" ]
  },
  "type": "bar"
}
""";

        var minifiedPayload = MarkdownVisualContract.CreatePayload(minified);
        var formattedPayload = MarkdownVisualContract.CreatePayload(formatted.Replace("\n", "\r\n"));

        Assert.Equal(minifiedPayload.Hash, formattedPayload.Hash);
        Assert.NotEqual(minifiedPayload.Base64, formattedPayload.Base64);
    }

    [Fact]
    public void MarkdownRenderer_Renders_Semantic_Chart_Blocks_By_SemanticKind_When_Language_Does_Not_Match_A_Renderer() {
        var md = """
```vendor-chart-json
{"type":"bar"}
```
""";
        var opts = new MarkdownRendererOptions();
        opts.Chart.Enabled = true;
        opts.ReaderOptions.FencedBlockExtensions.Add(new MarkdownFencedBlockExtension(
            "Vendor chart AST",
            new[] { "vendor-chart-json" },
            context => new SemanticFencedBlock(MarkdownSemanticKinds.Chart, context.Language, context.Content, context.Caption)));

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("class=\"omd-visual omd-chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"vendor-chart-json\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Can_Apply_Custom_Fenced_Code_Block_Renderers() {
        var md = "```ix-note\nhello <world>\n```";
        var opts = new MarkdownRendererOptions();
        opts.FencedCodeBlockRenderers.Add(new MarkdownFencedCodeBlockRenderer(
            "IX note",
            new[] { "ix-note" },
            (match, _) => $"<aside class=\"ix-note\" data-lang=\"{match.Language}\">{System.Net.WebUtility.HtmlEncode(match.RawContent)}</aside>") {
            BuildShellHeadHtml = (_, _) => "<style>.ix-note{border-left:4px solid #0a84ff;padding-left:12px;}</style>",
            BuildBeforeContentReplaceScript = _ => "window.__ixNoteBefore = (window.__ixNoteBefore || 0) + 1;",
            BuildAfterContentReplaceScript = _ => "window.__ixNoteAfter = (window.__ixNoteAfter || 0) + 1;"
        });

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);
        var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat", opts);

        Assert.Contains("class=\"ix-note\"", html, StringComparison.Ordinal);
        Assert.Contains("hello &lt;world&gt;", html, StringComparison.Ordinal);
        Assert.Contains(".ix-note", shell, StringComparison.Ordinal);
        Assert.Contains("__ixNoteBefore", shell, StringComparison.Ordinal);
        Assert.Contains("__ixNoteAfter", shell, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Converts_Generic_Dataview_Fences_To_Static_Table_Html() {
        var raw = "{\"title\":\"Replication Summary\",\"summary\":\"Latest replication posture\",\"kind\":\"generic_dataview_v1\",\"call_id\":\"call_123\",\"rows\":[[\"Server\",\"Fails\"],[\"AD0\",\"0\"],[\"AD1\",\"1\"]]}";
        var md = """
```dataview
{"title":"Replication Summary","summary":"Latest replication posture","kind":"generic_dataview_v1","call_id":"call_123","rows":[["Server","Fails"],["AD0","0"],["AD1","1"]]}
```
""";

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, MarkdownRendererPresets.CreateStrictMinimal());
        var payloadHash = MarkdownVisualContract.ComputePayloadHash(raw);

        Assert.Contains("class=\"omd-visual omd-dataview\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-contract=\"v1\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-kind=\"dataview\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"dataview\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-omd-visual-hash=\"{payloadHash}\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-title=\"Replication Summary\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-summary=\"Latest replication posture\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-kind=\"generic_dataview_v1\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-call-id=\"call_123\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-column-count=\"2\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-row-count=\"2\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-omd-dataview-payload-hash=\"{payloadHash}\"", html, StringComparison.Ordinal);
        Assert.Equal(raw, DecodeBase64Attribute(html, "data-omd-config-b64"));
        Assert.DoesNotContain("data-ix-title=", html, StringComparison.Ordinal);
        Assert.Contains("<caption>Replication Summary</caption>", html, StringComparison.Ordinal);
        Assert.Contains("<p class=\"omd-dataview-summary\">Latest replication posture</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Converts_Generic_Dataview_Fences_With_Neutral_Aliases_To_Static_Table_Html() {
        var raw = "{\"caption\":\"Replication Summary\",\"description\":\"Latest replication posture\",\"schema\":\"generic.dataview.v2\",\"callId\":\"call_456\",\"headers\":[\"Server\",\"Fails\"],\"items\":[{\"server\":\"AD0\",\"fails\":0},{\"server\":\"AD1\",\"fails\":1}]}";
        var md = """
```dataview
{"caption":"Replication Summary","description":"Latest replication posture","schema":"generic.dataview.v2","callId":"call_456","headers":["Server","Fails"],"items":[{"server":"AD0","fails":0},{"server":"AD1","fails":1}]}
```
""";

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, MarkdownRendererPresets.CreateStrictMinimal());
        var payloadHash = MarkdownVisualContract.ComputePayloadHash(raw);

        Assert.Contains("class=\"omd-visual omd-dataview\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-title=\"Replication Summary\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-summary=\"Latest replication posture\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-kind=\"generic.dataview.v2\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-call-id=\"call_456\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-column-count=\"2\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-row-count=\"2\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-omd-dataview-payload-hash=\"{payloadHash}\"", html, StringComparison.Ordinal);
        Assert.Equal(raw, DecodeBase64Attribute(html, "data-omd-config-b64"));
        Assert.Contains("<th>Server</th>", html, StringComparison.Ordinal);
        Assert.Contains("<th>Fails</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>AD0</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td>1</td>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-ix-title=", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Converts_IxDataview_Fences_To_Static_Table_Html() {
        var raw = "{\"title\":\"Replication Summary\",\"summary\":\"Latest replication posture\",\"kind\":\"ix_tool_dataview_v1\",\"call_id\":\"call_123\",\"rows\":[[\"Server\",\"Fails\"],[\"AD0\",\"0\"],[\"AD1\",\"1\"]]}";
        var md = """
```ix-dataview
{"title":"Replication Summary","summary":"Latest replication posture","kind":"ix_tool_dataview_v1","call_id":"call_123","rows":[["Server","Fails"],["AD0","0"],["AD1","1"]]}
```
""";

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal());
        var payloadHash = MarkdownVisualContract.ComputePayloadHash(raw);

        Assert.Contains("class=\"omd-visual omd-dataview\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-contract=\"v1\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-kind=\"dataview\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-fence-language=\"ix-dataview\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-omd-visual-hash=\"{payloadHash}\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-config-format=\"json\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-config-encoding=\"base64-utf8\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-config-b64=\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-dataview-table\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-title=\"Replication Summary\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-summary=\"Latest replication posture\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-kind=\"ix_tool_dataview_v1\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-call-id=\"call_123\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-column-count=\"2\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-row-count=\"2\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-omd-dataview-payload-hash=\"{payloadHash}\"", html, StringComparison.Ordinal);
        Assert.Equal(raw, DecodeBase64Attribute(html, "data-omd-config-b64"));
        Assert.Contains("data-ix-title=\"Replication Summary\"", html, StringComparison.Ordinal);
        Assert.Contains("data-ix-summary=\"Latest replication posture\"", html, StringComparison.Ordinal);
        Assert.Contains("data-ix-kind=\"ix_tool_dataview_v1\"", html, StringComparison.Ordinal);
        Assert.Contains("data-ix-call-id=\"call_123\"", html, StringComparison.Ordinal);
        Assert.Contains("data-ix-column-count=\"2\"", html, StringComparison.Ordinal);
        Assert.Contains("data-ix-row-count=\"2\"", html, StringComparison.Ordinal);
        Assert.Contains("<caption>Replication Summary</caption>", html, StringComparison.Ordinal);
        Assert.Contains("<p class=\"omd-dataview-summary\">Latest replication posture</p>", html, StringComparison.Ordinal);
        Assert.Contains($"data-ix-payload-hash=\"{payloadHash}\"", html, StringComparison.Ordinal);
        Assert.Contains("<th>Server</th>", html, StringComparison.Ordinal);
        Assert.Contains("<th>Fails</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>AD0</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td>1</td>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Converts_IxDataview_Columns_And_Object_Records_To_Static_Table_Html() {
        var md = """
```ix-dataview
{"kind":"ix_tool_dataview_v1","columns":["Server","Fails"],"records":[{"Server":"AD0","Fails":0},{"Server":"AD1","Fails":1}]}
```
""";

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal());

        Assert.Contains("class=\"omd-visual omd-dataview\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-kind=\"dataview\"", html, StringComparison.Ordinal);
        Assert.Contains("<th>Server</th>", html, StringComparison.Ordinal);
        Assert.Contains("<th>Fails</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>AD0</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td>1</td>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Converts_Dataview_Rows_Object_Payloads_With_CaseInsensitive_Columns() {
        var md = """
```dataview
{"headers":["Server","Fails"],"rows":[{"server":"AD0","fails":0},{"server":"AD1","fails":1}]}
```
""";

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, MarkdownRendererPresets.CreateStrictMinimal());

        Assert.Contains("class=\"omd-visual omd-dataview\"", html, StringComparison.Ordinal);
        Assert.Contains("<th>Server</th>", html, StringComparison.Ordinal);
        Assert.Contains("<th>Fails</th>", html, StringComparison.Ordinal);
        Assert.Contains("<td>AD0</td>", html, StringComparison.Ordinal);
        Assert.Contains("<td>1</td>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Leaves_Invalid_IxDataview_Fences_As_Code_Blocks() {
        var md = """
```ix-dataview
{ not json
```
""";

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal());

        Assert.DoesNotContain("class=\"omd-dataview\"", html, StringComparison.Ordinal);
        Assert.Contains("<pre><code", html, StringComparison.Ordinal);
        Assert.Contains("{ not json", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Custom_Fenced_Code_Block_Renderers_Can_Override_BuiltIn_Aliases() {
        var md = "```chart\n{\"type\":\"bar\"}\n```";
        var opts = new MarkdownRendererOptions();
        opts.Chart.Enabled = true;
        opts.FencedCodeBlockRenderers.Add(new MarkdownFencedCodeBlockRenderer(
            "Chart override",
            new[] { "chart" },
            (_, _) => "<div class=\"custom-chart\">override</div>"));

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("class=\"custom-chart\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"omd-chart\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Shell_Contains_ChartJs_When_Enabled() {
        var opts = new MarkdownRendererOptions();
        opts.Chart.Enabled = true;

        var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat", opts);
        Assert.Contains("chart.umd", shell, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("data-omd-visual-rendered", shell, StringComparison.Ordinal);
        Assert.Contains("data-chart-rendered", shell, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Shell_Contains_VisNetwork_When_Enabled() {
        var opts = new MarkdownRendererOptions();
        opts.Network.Enabled = true;

        var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat", opts);
        Assert.Contains("vis-network.min.js", shell, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("vis-network.min.css", shell, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(".omd-network-canvas", shell, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-rendered", shell, StringComparison.Ordinal);
        Assert.Contains("data-network-rendered", shell, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_BaseHref_Is_Emitted_As_Base_Tag() {
        var opts = new MarkdownRendererOptions { BaseHref = "https://example.com/" };
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("[link](/x)", opts);
        Assert.Contains("<base href=\"https://example.com/\">", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Defaults_To_Hardened_External_Links() {
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("[x](https://example.com)");
        Assert.Contains("target=\"_blank\"", html, StringComparison.Ordinal);
        Assert.Contains("rel=\"", html, StringComparison.Ordinal);
        Assert.Contains("noopener", html, StringComparison.Ordinal);
        Assert.Contains("noreferrer", html, StringComparison.Ordinal);
        Assert.Contains("nofollow", html, StringComparison.Ordinal);
        Assert.Contains("ugc", html, StringComparison.Ordinal);
        Assert.Contains("referrerpolicy=\"no-referrer\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Can_Restrict_Absolute_Http_Links_To_Base_Origin() {
        var opts = new MarkdownRendererOptions { BaseHref = "https://example.com/" };
        opts.HtmlOptions.RestrictHttpLinksToBaseOrigin = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("[x](https://other.example.com/path)", opts);
        Assert.DoesNotContain("other.example.com", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Can_Restrict_Absolute_Http_Images_To_Base_Origin() {
        var opts = new MarkdownRendererOptions { BaseHref = "https://example.com/" };
        opts.HtmlOptions.RestrictHttpImagesToBaseOrigin = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("![alt](https://other.example.com/a.png)", opts);
        Assert.DoesNotContain("<img", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("omd-image-blocked", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Defaults_Emit_Image_Hardening_Attributes() {
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("![alt](/a.png)");
        Assert.Contains("<img", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("loading=\"lazy\"", html, StringComparison.Ordinal);
        Assert.Contains("decoding=\"async\"", html, StringComparison.Ordinal);
        Assert.Contains("referrerpolicy=\"no-referrer\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Defaults_Block_External_Absolute_Http_Images() {
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("![alt](https://example.com/a.png)");
        Assert.DoesNotContain("<img", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("omd-image-blocked", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Converts_Fenced_Math_To_Display_Math_When_Enabled() {
        string md = """
```math
x^2 + 1
```
""";
        var opts = new MarkdownRendererOptions { BaseHref = "https://example.com/" };
        opts.HtmlOptions.BlockExternalHttpImages = false;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);
        Assert.Contains("class=\"omd-math\"", html, StringComparison.Ordinal);
        Assert.Contains("$$", html, StringComparison.Ordinal);
        Assert.Contains("x^2 + 1", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Can_Apply_Custom_Html_PostProcessors() {
        var opts = new MarkdownRendererOptions();
        opts.HtmlPostProcessors.Add((html, _) => html + "<div id=\"post\">x</div>");

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("hello", opts);
        Assert.Contains("id=\"post\"", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_SoftWrapped_Strong_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeSoftWrappedStrongSpans = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("**Status\nHEALTHY**", opts);
        Assert.Contains("Status HEALTHY", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("Status\nHEALTHY", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_InlineCode_LineBreaks_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeInlineCodeSpanLineBreaks = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("`line1\nline2`", opts);
        Assert.Contains("line1 line2", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("line1\nline2", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_NestedStrongDelimiters_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeNestedStrongDelimiters = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("- Signal **Current comparison used **System** log only.**", opts);
        Assert.Contains("<strong>Current comparison used System log only.</strong>", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("used **System** log only.**", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_NestedStrongDelimiters_InsideLabeledBullets_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeNestedStrongDelimiters = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("- Why it matters **Current comparison used **System** log only.**", opts);
        Assert.Contains("Why it matters <strong>Current comparison used System log only.</strong>", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("used **System** log only.**", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_NestedStrongDelimiters_WithoutTouchingInlineCode_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeNestedStrongDelimiters = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(
            "- Signal **pattern `a**b` seen, mostly from **Service Control Manager**.**",
            opts);

        Assert.Contains("<code>a**b</code>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("mostly from Service Control Manager.", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("from **Service", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_RepeatedStrongDelimiterRuns_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeLooseStrongDelimiters = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("- Overall health ****healthy****", opts);
        Assert.Contains("<strong>healthy</strong>", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("****healthy****", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_TightArrowAndColonSpacing_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeTightStrongBoundaries = true,
            NormalizeTightArrowStrongBoundaries = true,
            NormalizeTightColonSpacing = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("- Signal **Healthy baseline exists now** ->**Why it matters:**missing coverage", opts);
        Assert.Contains("-&gt; <strong>Why it matters:</strong> missing coverage", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_SignalFlowLabelSpacing_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeSignalFlowLabelSpacing = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(
            "- Signal -> Why it matters:missing coverage -> Next action:review defaults",
            opts);
        Assert.Contains("Why it matters: missing coverage", htmlOut, StringComparison.Ordinal);
        Assert.Contains("Next action: review defaults", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_BrokenStrongArrowLabels_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeBrokenStrongArrowLabels = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("- Signal **No current failures -> **Why it matters:** transport/auth issues", opts);
        Assert.Contains("<strong>No current failures</strong> -&gt; <strong>Why it matters:</strong> transport/auth issues", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_CompactHeadingAndStrongLabelListBoundaries_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeHeadingListBoundaries = true,
            NormalizeCompactStrongLabelListBoundaries = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("## Wynik ogólny- **Replication:** wcześniej zdrowa ✅- **FSMO:** technicznie OK", opts);
        Assert.Contains("<h2", htmlOut, StringComparison.Ordinal);
        Assert.Contains("Wynik og", htmlOut, StringComparison.Ordinal);
        Assert.Equal(2, Count(htmlOut, "<li"));
        Assert.Contains("<strong>Replication:</strong>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>FSMO:</strong>", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_CompactHeadingBoundaries_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeCompactHeadingBoundaries = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("previous shutdown was unexpected### Reason", opts);
        Assert.Contains("<p>previous shutdown was unexpected</p>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<h3", htmlOut, StringComparison.Ordinal);
        Assert.Contains("Reason", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_ColonListBoundaries_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeColonListBoundaries = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("Następny najlepszy krok:- **`ad_domain_controller_facts`**", opts);
        Assert.Contains("<p>Następny najlepszy krok:</p>", htmlOut, StringComparison.Ordinal);
        Assert.Equal(1, Count(htmlOut, "<li"));
        Assert.Contains("ad_domain_controller_facts", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_CompactMermaidFenceBodyBoundary_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };
        opts.Mermaid.Enabled = true;

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("```mermaidflowchart LR A-->B\n```", opts);

        Assert.Contains("class=\"mermaid\"", htmlOut, StringComparison.Ordinal);
        Assert.Contains("flowchart LR A--&gt;B", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Normalizes_OrderedListParenCaretAndParentheticalSpacing_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeOrderedListParenMarkers = true,
            NormalizeOrderedListCaretArtifacts = true,
            NormalizeTightParentheticalSpacing = true
        };

        var markdown = """
1) First check
2.^ **Delegation risk audit**
3. **Deleted object remnants**(SID left in ACL path)
""";

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, opts);
        Assert.Equal(3, Count(htmlOut, "<li"));
        Assert.Contains("<strong>Delegation risk audit</strong>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>Deleted object remnants</strong> (SID left in ACL path)", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_DoesNot_Normalize_TightParentheticalSpacing_InsideInlineCode() {
        var opts = new MarkdownRendererOptions {
            NormalizeTightParentheticalSpacing = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("Use `Get-ADUser(SIDHistory)` and **Deleted object remnants**(SID left in ACL path)", opts);
        Assert.Contains("<code>Get-ADUser(SIDHistory)</code>", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("<code>Get-ADUser (SIDHistory)</code>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>Deleted object remnants</strong> (SID left in ACL path)", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Can_Apply_Markdown_PreProcessors() {
        var opts = new MarkdownRendererOptions();
        opts.MarkdownPreProcessors.Add((markdown, _) => markdown.Replace("{{name}}", "IntelligenceX"));

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("hello {{name}}", opts);
        Assert.Contains("hello IntelligenceX", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_MarkdownPreProcessors_Run_After_PreParse_Normalization() {
        var opts = new MarkdownRendererOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };
        opts.Mermaid.Enabled = true;
        opts.MarkdownPreProcessors.Add((markdown, _) =>
            markdown.Replace("```mermaid\nflowchart LR", "```mermaid\ngraph TD"));

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("```mermaidflowchart LR A-->B\n```", opts);

        Assert.Contains("class=\"mermaid\"", htmlOut, StringComparison.Ordinal);
        Assert.Contains("graph TD A--&gt;B", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("flowchart LR A--&gt;B", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRendererPreProcessorPipeline_Applies_PreProcessors_In_Order() {
        var opts = new MarkdownRendererOptions();
        opts.MarkdownPreProcessors.Add((markdown, _) => markdown.Replace("{{name}}", "IntelligenceX"));
        opts.MarkdownPreProcessors.Add((markdown, _) => markdown.Replace("hello IntelligenceX", "hello OfficeIMO"));

        var processed = MarkdownRendererPreProcessorPipeline.Apply("hello {{name}}", opts);

        Assert.Equal("hello OfficeIMO", processed);
    }

    [Fact]
    public void MarkdownRenderer_Preserves_PreParse_Normalization_From_ReaderInputNormalization() {
        var opts = new MarkdownRendererOptions {
            ReaderOptions = new MarkdownReaderOptions {
                InputNormalization = new MarkdownInputNormalizationOptions {
                    NormalizeOrderedListParenMarkers = true
                },
                HtmlBlocks = false,
                InlineHtml = false
            }
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("1) First check", opts);

        Assert.Contains("<ol>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<li>First check</li>", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_IntelligenceXTranscriptPreset_Enables_Text_Normalization() {
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("**Status\nHEALTHY**\n\n`a\nb`\n\nUse \\`/act act_001\\`.\n\nStatus **Healthy**next\n\ncheck ** LDAP/Kerberos health on all DCs** next\n\n- Signal **Current comparison used **System** log only.**\n- Signal **Healthy baseline exists now** ->**Why it matters:**missing coverage\n- Signal **No current failures -> **Why it matters:** transport/auth issues\n\n## Wynik ogólny- **Replication:** wcześniej zdrowa ✅- **FSMO:** technicznie OK\n\nprevious shutdown was unexpected### Reason\n\nNastępny najlepszy krok:- **`ad_domain_controller_facts`**\n\n1) First check\n2.^ **Delegation risk audit**\n3. **Deleted object remnants**(SID left in ACL path)\n\nCommand: `Get-ADUser(SIDHistory)`", opts);

        Assert.Contains("Status HEALTHY", htmlOut, StringComparison.Ordinal);
        Assert.Contains("a b", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<code>/act act_001</code>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>Healthy</strong> next", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>LDAP/Kerberos health on all DCs</strong> next", htmlOut, StringComparison.Ordinal);
        Assert.Contains("Current comparison used System log only.", htmlOut, StringComparison.Ordinal);
        Assert.Contains("-&gt; <strong>Why it matters:</strong> missing coverage", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>No current failures</strong> -&gt; <strong>Why it matters:</strong> transport/auth issues", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<h2", htmlOut, StringComparison.Ordinal);
        Assert.Contains("Wynik og", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>Replication:</strong>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>FSMO:</strong>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<h3", htmlOut, StringComparison.Ordinal);
        Assert.Contains("Reason", htmlOut, StringComparison.Ordinal);
        Assert.Contains("Następny najlepszy krok:", htmlOut, StringComparison.Ordinal);
        Assert.Contains("ad_domain_controller_facts", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<li>First check</li>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>Delegation risk audit</strong>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>Deleted object remnants</strong> (SID left in ACL path)", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<code>Get-ADUser(SIDHistory)</code>", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("<code>Get-ADUser (SIDHistory)</code>", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("used **System** log only.**", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_IntelligenceXTranscriptPreset_Normalizes_CompactMermaidFenceBodies() {
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscript();

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("```mermaidflowchart LR A-->B\n```", opts);

        Assert.Contains("class=\"mermaid\"", htmlOut, StringComparison.Ordinal);
        Assert.Contains("flowchart LR", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_IntelligenceXTranscriptPreset_Normalizes_OrderedListMarkerSpacing() {
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        var markdown = """
1. **Privilege hygiene sweep**(Domain Admins + other privileged groups)
2.** Delegation risk audit**(unconstrained / constrained / protocol transition)
3.** Replication + DC health snapshot** (stale links, failing partners)
""";

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, opts);
        Assert.Equal(3, Count(htmlOut, "<li"));
        Assert.DoesNotContain("2.**", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("3.**", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_IntelligenceXTranscriptPreset_Normalizes_CollapsedOrderedListTranscriptArtifacts() {
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        var markdown = "1) **Privilege hygiene sweep(Domain Admins + other privileged groups, nested exposure) 2)** Delegation risk audit**(unconstrained / constrained / protocol transition) 3)** Replication + DC health snapshot** (stale links, failing partners, LDAP/Kerberos basics)";

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, opts);

        Assert.Equal(3, Count(htmlOut, "<li"));
        Assert.Contains("<strong>Privilege hygiene sweep</strong> (Domain Admins + other privileged groups, nested exposure)", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>Delegation risk audit</strong> (unconstrained / constrained / protocol transition)", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<strong>Replication + DC health snapshot</strong> (stale links, failing partners, LDAP/Kerberos basics)", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_IntelligenceXTranscriptPreset_DoesNotCollapse_Adjacent_UnorderedListItems_WithStrongLabels() {
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimal();
        var markdown = """
- **AD1:** 875 Events  
- **AD2:** 353 Events

Top-IDs:
- **AD1:** `7034` (666), `7023` (97), `10010` (95)
- **AD2:** `1801` (162), `36874` (67)
""";

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, opts);
        Assert.Equal(4, Count(htmlOut, "<li"));
        Assert.DoesNotContain("-** AD2:**", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("<dl>", htmlOut, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void MarkdownRenderer_IntelligenceXTranscriptMinimalPortable_Disables_Callouts_TaskLists_And_LiteralAutolinks() {
        var opts = MarkdownRendererPresets.CreateIntelligenceXTranscriptMinimalPortable();
        var markdown = """
> [!NOTE]
> body

- [ ] task

Visit https://example.com now.
""";

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, opts);
        Assert.DoesNotContain("class=\"callout", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("contains-task-list", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("task-list-item-checkbox", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("href=\"https://example.com\"", htmlOut, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("[!NOTE]", htmlOut, StringComparison.Ordinal);
        Assert.Contains("[ ] task", htmlOut, StringComparison.Ordinal);
        Assert.Contains("Visit https://example.com now.", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Allows_SameOrigin_Absolute_Http_Images_When_BaseHref_Is_Set() {
        var opts = new MarkdownRendererOptions { BaseHref = "https://example.com/" };
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("![alt](https://example.com/a.png)", opts);
        Assert.Contains("<img", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("loading=\"lazy\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlOptions_Can_Restrict_Absolute_Http_Links_By_Host_AllowList() {
        var opts = new MarkdownRendererOptions();
        opts.HtmlOptions.AllowedHttpLinkHosts.Add("example.com");

        var ok = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("[x](https://example.com/a)", opts);
        Assert.Contains("href=\"https://example.com/a\"", ok, StringComparison.OrdinalIgnoreCase);

        var blocked = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("[x](https://evil.example/a)", opts);
        Assert.DoesNotContain("evil.example", blocked, StringComparison.OrdinalIgnoreCase);
        Assert.Contains(">x<", blocked, StringComparison.Ordinal);
    }

    [Fact]
    public void HtmlOptions_Can_Restrict_Absolute_Http_Images_By_Host_AllowList() {
        var opts = new MarkdownRendererOptions();
        opts.HtmlOptions.BlockExternalHttpImages = false; // isolate host allowlist behavior
        opts.HtmlOptions.AllowedHttpImageHosts.Add(".example.com"); // apex + subdomains

        var ok = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("![a](https://a.example.com/x.png)", opts);
        Assert.Contains("<img", ok, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("a.example.com", ok, StringComparison.OrdinalIgnoreCase);

        var blocked = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("![a](https://evil.test/x.png)", opts);
        Assert.DoesNotContain("<img", blocked, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("omd-image-blocked", blocked, StringComparison.Ordinal);
    }

    [Fact]
    public void Task_Lists_Emit_GithubLike_Classes() {
        var md = """
- [ ] Todo
- [x] Done
""";
        var doc = MarkdownReader.Parse(md);
        var html = doc.ToHtmlFragment(new HtmlOptions { Style = HtmlStyle.Plain, CssDelivery = CssDelivery.None, BodyClass = null });

        Assert.Contains("contains-task-list", html, StringComparison.Ordinal);
        Assert.Contains("task-list-item", html, StringComparison.Ordinal);
        Assert.Contains("task-list-item-checkbox", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Preserves_PreferNarrativeSingleLineDefinitions_From_ReaderOptions() {
        var markdown = """
Status: Healthy

Next paragraph.
""";
        var opts = new MarkdownRendererOptions {
            ReaderOptions = new MarkdownReaderOptions {
                PreferNarrativeSingleLineDefinitions = true,
                HtmlBlocks = false,
                InlineHtml = false
            }
        };

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, opts);

        Assert.DoesNotContain("<dl>", html, StringComparison.OrdinalIgnoreCase);
        Assert.Contains("<p>Status: Healthy</p>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Next paragraph.</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Preserves_TaskLists_Flag_From_ReaderOptions() {
        var markdown = """
- [ ] Todo
- [x] Done
""";
        var opts = new MarkdownRendererOptions {
            ReaderOptions = new MarkdownReaderOptions {
                TaskLists = false,
                HtmlBlocks = false,
                InlineHtml = false
            }
        };

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, opts);

        Assert.DoesNotContain("contains-task-list", html, StringComparison.Ordinal);
        Assert.DoesNotContain("task-list-item-checkbox", html, StringComparison.Ordinal);
        Assert.Contains("[ ] Todo", html, StringComparison.Ordinal);
        Assert.Contains("[x] Done", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Preserves_TocPlaceholders_And_Footnotes_Flags_From_ReaderOptions() {
        var markdown = """
[TOC]

Lead[^1]

[^1]: Footnote text
""";
        var readerOptions = MarkdownReaderOptions.CreateCommonMarkProfile();
        readerOptions.HtmlBlocks = false;
        readerOptions.InlineHtml = false;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(markdown, new MarkdownRendererOptions {
            ReaderOptions = readerOptions
        });

        Assert.Contains("<p>[TOC]</p>", html, StringComparison.Ordinal);
        Assert.Contains("<p>Lead[^1]</p>", html, StringComparison.Ordinal);
        Assert.Contains("<p>[^1]: Footnote text</p>", html, StringComparison.Ordinal);
        Assert.DoesNotContain("footnotes", html, StringComparison.OrdinalIgnoreCase);
    }
    private static int Count(string value, string token) {
        if (string.IsNullOrEmpty(value) || string.IsNullOrEmpty(token)) return 0;

        int index = 0;
        int count = 0;
        while ((index = value.IndexOf(token, index, StringComparison.Ordinal)) >= 0) {
            count++;
            index += token.Length;
        }

        return count;
    }

    private static string DecodeBase64Attribute(string html, string attributeName) {
        var marker = attributeName + "=\"";
        var start = html.IndexOf(marker, StringComparison.Ordinal);
        Assert.True(start >= 0, $"Expected attribute {attributeName} in HTML.");

        start += marker.Length;
        var end = html.IndexOf('"', start);
        Assert.True(end >= start, $"Expected closing quote for attribute {attributeName}.");

        var encoded = html.Substring(start, end - start);
        var bytes = Convert.FromBase64String(System.Net.WebUtility.HtmlDecode(encoded));
        return Encoding.UTF8.GetString(bytes).TrimEnd('\r', '\n');
    }

}

