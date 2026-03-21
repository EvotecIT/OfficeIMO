using System.Globalization;
using System.Linq;
using System.Text;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;
using OfficeIMO.MarkdownRenderer;
using OfficeIMO.MarkdownRenderer.SamplePlugin;
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
    public void MarkdownRenderer_Chart_Emits_Shared_Visual_Title_From_Fence_Metadata() {
        var configJson = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var md = $"""
```chart title="Quarterly Overview"
{configJson}
```
""";
        var opts = new MarkdownRendererOptions();
        opts.Chart.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("data-omd-fence-language=\"chart\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-title=\"Quarterly Overview\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Chart_Honors_Brace_Style_Fence_Id_And_Classes() {
        var configJson = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var md = "```chart {#quarterly-overview .wide .accent title=\"Quarterly Overview\" pinned}\n"
            + configJson
            + "\n```";
        var opts = new MarkdownRendererOptions();
        opts.Chart.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("data-omd-fence-info=\"{#quarterly-overview .wide .accent title=&quot;Quarterly Overview&quot; pinned}\"", html, StringComparison.Ordinal);
        Assert.Contains("id=\"quarterly-overview\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-chart wide accent\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-title=\"Quarterly Overview\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Chart_Does_Not_Partially_Apply_Malformed_Brace_Metadata() {
        var configJson = "{\"type\":\"bar\",\"data\":{\"labels\":[\"A\"],\"datasets\":[{\"label\":\"Count\",\"data\":[1]}]}}";
        var md = "```chart {#quarterly-overview .wide title=\"Quarterly Overview\"\n"
            + configJson
            + "\n```";
        var opts = new MarkdownRendererOptions();
        opts.Chart.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("data-omd-fence-info=\"{#quarterly-overview .wide title=&quot;Quarterly Overview&quot;\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"quarterly-overview\"", html, StringComparison.Ordinal);
        Assert.DoesNotContain("class=\"omd-visual omd-chart wide", html, StringComparison.Ordinal);
        Assert.DoesNotContain("data-omd-visual-title=\"Quarterly Overview\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-chart\"", html, StringComparison.Ordinal);
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
    public void MarkdownRenderer_Network_Emits_Shared_Visual_Title_From_Fence_Metadata() {
        var md = """
```network title="Relationship Map"
{"nodes":[{"id":"A","label":"User"},{"id":"B","label":"Group"}],"edges":[{"from":"A","to":"B","label":"memberOf"}]}
```
""";
        var opts = new MarkdownRendererOptions();
        opts.Network.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("data-omd-fence-language=\"network\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-title=\"Relationship Map\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Network_Honors_Brace_Style_Fence_Id_And_Classes() {
        var md = """
```network {#relationship-map .wide .interactive title="Relationship Map"}
{"nodes":[{"id":"A","label":"User"},{"id":"B","label":"Group"}],"edges":[{"from":"A","to":"B","label":"memberOf"}]}
```
""";
        var opts = new MarkdownRendererOptions();
        opts.Network.Enabled = true;

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("id=\"relationship-map\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-network wide interactive\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-title=\"Relationship Map\"", html, StringComparison.Ordinal);
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
    public void MarkdownVisualContract_Can_Apply_Fence_Metadata_To_Host_Elements() {
        var raw = "{\"type\":\"bar\"}";
        var payload = MarkdownVisualContract.CreatePayload(raw);
        var fenceInfo = MarkdownCodeFenceInfo.Parse("vendor-chart {#sales-summary .wide .accent title=\"Quarterly Overview\" pinned}");
        var html = MarkdownVisualContract.BuildElementHtml(
            "div",
            "omd-visual omd-custom",
            "custom-chart",
            "vendor-chart",
            payload,
            fenceInfo,
            new KeyValuePair<string, string?>("data-custom-hash", payload.Hash));

        Assert.Contains("data-omd-fence-info=\"{#sales-summary .wide .accent title=&quot;Quarterly Overview&quot; pinned}\"", html, StringComparison.Ordinal);
        Assert.Contains("id=\"sales-summary\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-custom wide accent\"", html, StringComparison.Ordinal);
        Assert.Contains($"data-custom-hash=\"{payload.Hash}\"", html, StringComparison.Ordinal);
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
    public void MarkdownRendererOptions_Can_Register_And_Parse_Fence_Option_Schemas() {
        var schema = new MarkdownFenceOptionSchema(
            "vendor.visual-options",
            "Vendor Visual Options",
            new[] { "vendor-chart" },
            new[] {
                MarkdownFenceOptionDefinition.Boolean("pinned"),
                MarkdownFenceOptionDefinition.Int32(
                    "maxItems",
                    aliases: new[] { "limit" },
                    validator: rawValue => int.TryParse(rawValue, System.Globalization.NumberStyles.Integer, System.Globalization.CultureInfo.InvariantCulture, out var parsed)
                        && parsed > 0
                        ? null
                        : "Expected a positive integer value."),
                MarkdownFenceOptionDefinition.String("theme")
            });

        var options = new MarkdownRendererOptions();
        options.ApplyFenceOptionSchema(schema);

        Assert.True(options.HasFenceOptionSchema(schema));
        Assert.True(options.TryGetFenceOptionSchema("vendor-chart", out var resolvedSchema));
        Assert.Equal(schema.Id, resolvedSchema.Id);

        var fenceInfo = MarkdownCodeFenceInfo.Parse("vendor-chart title=\"Quarterly Revenue\" pinned limit=12 theme=\"sunset\" custom=true");
        Assert.True(options.TryParseFenceOptions("vendor-chart", fenceInfo, out var parsed));
        Assert.True(parsed.IsValid);
        Assert.True(parsed.TryGetBoolean("pinned", out var pinned));
        Assert.True(pinned);
        Assert.True(parsed.TryGetInt32("maxItems", out var maxItems));
        Assert.Equal(12, maxItems);
        Assert.True(parsed.TryGetString("theme", out var theme));
        Assert.Equal("sunset", theme);
        Assert.Contains("custom", parsed.UnknownOptions);
        Assert.DoesNotContain("title", parsed.UnknownOptions);
    }

    [Fact]
    public void MarkdownRendererPlugin_Can_Carry_Fence_Option_Schemas() {
        var schema = new MarkdownFenceOptionSchema(
            "vendor.visual-options",
            "Vendor Visual Options",
            new[] { "vendor-chart" },
            new[] {
                MarkdownFenceOptionDefinition.Boolean("pinned"),
                MarkdownFenceOptionDefinition.Int32("maxItems", aliases: new[] { "limit" })
            });

        var plugin = new MarkdownRendererPlugin(
            "Vendor Visuals",
            new Func<MarkdownFencedCodeBlockRenderer>[] {
                () => new MarkdownFencedCodeBlockRenderer(
                    "Vendor chart",
                    new[] { "vendor-chart" },
                    (_, _) => "<div class=\"vendor-chart\"></div>")
            },
            new[] { schema });

        var options = new MarkdownRendererOptions();
        options.ApplyPlugin(plugin);

        Assert.True(options.HasPlugin(plugin));
        Assert.True(options.HasFenceOptionSchema(schema));
        Assert.True(options.TryParseFenceOptions("vendor-chart", MarkdownCodeFenceInfo.Parse("vendor-chart pinned limit=5"), out var parsed));
        Assert.True(parsed.TryGetBoolean("pinned", out var pinned));
        Assert.True(pinned);
        Assert.True(parsed.TryGetInt32("maxItems", out var maxItems));
        Assert.Equal(5, maxItems);
    }

    [Fact]
    public void MarkdownRendererPlugin_Can_Carry_Reader_Configuration_And_Remain_Idempotent() {
        var plugin = new MarkdownRendererPlugin(
            "Vendor Transcript Visuals",
            new Func<MarkdownFencedCodeBlockRenderer>[] {
                () => new MarkdownFencedCodeBlockRenderer(
                    "Vendor chart",
                    new[] { "vendor-chart" },
                    (_, _) => "<div class=\"vendor-chart\"></div>")
            },
            apply: options => {
                options.ReaderOptions.PreferNarrativeSingleLineDefinitions = true;
                if (!options.ReaderOptions.DocumentTransforms.Any(transform => transform is MarkdownSimpleDefinitionListParagraphTransform)) {
                    options.ReaderOptions.DocumentTransforms.Add(new MarkdownSimpleDefinitionListParagraphTransform());
                }
            });

        var options = new MarkdownRendererOptions();
        options.ApplyPlugin(plugin);
        options.ApplyPlugin(plugin);

        Assert.True(options.HasPlugin(plugin));
        Assert.True(options.ReaderOptions.PreferNarrativeSingleLineDefinitions);
        Assert.Equal(1, options.FencedCodeBlockRenderers.Count(renderer => renderer.Languages.Contains("vendor-chart", StringComparer.OrdinalIgnoreCase)));
        Assert.Equal(1, options.ReaderOptions.DocumentTransforms.Count(transform => transform is MarkdownSimpleDefinitionListParagraphTransform));
    }

    [Fact]
    public void MarkdownRendererPlugin_And_FeaturePack_Can_Carry_Visual_RoundTrip_Hints() {
        var hint = new MarkdownVisualElementRoundTripHint(
            "vendor.caption",
            "Vendor caption",
            context => context.CreateBlock(caption: "Caption"));
        var readerTransform = new MarkdownInlineNormalizationTransform(new MarkdownInputNormalizationOptions {
            NormalizeTightColonSpacing = true
        });
        var transform = new MarkdownJsonVisualCodeBlockTransform(MarkdownVisualFenceLanguageMode.GenericSemanticFence);
        var rendererTransform = new RendererAppendParagraphTransform("renderer tail");
        var elementConverter = new HtmlElementBlockConverter(
            "vendor.custom-html",
            "Vendor custom HTML",
            _ => Array.Empty<IMarkdownBlock>());
        var inlineConverter = new HtmlInlineElementConverter(
            "vendor.inline-html",
            "Vendor inline HTML",
            _ => Array.Empty<IMarkdownInline>());
        var plugin = new MarkdownRendererPlugin(
            "Vendor Visuals",
            new Func<MarkdownFencedCodeBlockRenderer>[] {
                () => new MarkdownFencedCodeBlockRenderer(
                    "Vendor chart",
                    new[] { "vendor-chart" },
                    (_, _) => "<div class=\"vendor-chart\"></div>")
            },
            readerDocumentTransforms: new[] { readerTransform },
            htmlDocumentTransforms: new[] { transform },
            rendererDocumentTransforms: new[] { rendererTransform },
            htmlElementBlockConverters: new[] { elementConverter },
            htmlInlineElementConverters: new[] { inlineConverter },
            visualElementRoundTripHints: new[] { hint });
        var featurePack = new MarkdownRendererFeaturePack(
            "vendor.visual-pack",
            "Vendor Visual Pack",
            new[] { plugin });

        Assert.Single(plugin.ReaderDocumentTransforms);
        Assert.Same(readerTransform, plugin.ReaderDocumentTransforms[0]);
        Assert.Single(plugin.HtmlDocumentTransforms);
        Assert.Same(transform, plugin.HtmlDocumentTransforms[0]);
        Assert.Single(plugin.RendererDocumentTransforms);
        Assert.Same(rendererTransform, plugin.RendererDocumentTransforms[0]);
        Assert.Single(plugin.HtmlElementBlockConverters);
        Assert.Same(elementConverter, plugin.HtmlElementBlockConverters[0]);
        Assert.Single(plugin.HtmlInlineElementConverters);
        Assert.Same(inlineConverter, plugin.HtmlInlineElementConverters[0]);
        Assert.Single(plugin.VisualElementRoundTripHints);
        Assert.Equal("vendor.caption", plugin.VisualElementRoundTripHints[0].Id);
        Assert.Single(featurePack.ReaderDocumentTransforms);
        Assert.Same(readerTransform, featurePack.ReaderDocumentTransforms[0]);
        Assert.Single(featurePack.HtmlDocumentTransforms);
        Assert.Same(transform, featurePack.HtmlDocumentTransforms[0]);
        Assert.Single(featurePack.RendererDocumentTransforms);
        Assert.Same(rendererTransform, featurePack.RendererDocumentTransforms[0]);
        Assert.Single(featurePack.HtmlElementBlockConverters);
        Assert.Same(elementConverter, featurePack.HtmlElementBlockConverters[0]);
        Assert.Single(featurePack.HtmlInlineElementConverters);
        Assert.Same(inlineConverter, featurePack.HtmlInlineElementConverters[0]);
        Assert.Single(featurePack.VisualElementRoundTripHints);
        Assert.Equal("vendor.caption", featurePack.VisualElementRoundTripHints[0].Id);
    }

    [Fact]
    public void SampleMarkdownRenderer_StatusPanelPlugin_Can_Render_Shared_Visual_Host_Html() {
        const string raw = """
{"title":"Operations Overview","summary":"All checks passing","status":"healthy","caption":"Panel caption"}
""";
        var options = MarkdownRendererPresets.CreateStrictMinimal();

        SampleMarkdownRenderer.ApplyStatusPanels(options);
        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("```status-panel\n" + raw + "\n```", options);

        Assert.True(SampleMarkdownRenderer.HasStatusPanels(options));
        Assert.Single(SampleMarkdownRenderer.StatusPanelPlugin.VisualElementRoundTripHints);
        Assert.Contains("class=\"omd-visual omd-status-panel\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-visual-kind=\"status-panel\"", html, StringComparison.Ordinal);
        Assert.Contains("data-sample-panel-caption=\"Panel caption\"", html, StringComparison.Ordinal);
        Assert.Contains("Operations Overview", html, StringComparison.Ordinal);
        Assert.Contains("All checks passing", html, StringComparison.Ordinal);
    }

    [Fact]
    public void SampleMarkdownRenderer_StatusPanelFeaturePack_Carries_Renderer_And_RoundTrip_Contracts() {
        var options = MarkdownRendererPresets.CreateStrictMinimal();

        options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);
        options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);

        Assert.True(SampleMarkdownRenderer.HasStatusPanelFeaturePack(options));
        Assert.True(options.HasPlugin(SampleMarkdownRenderer.StatusPanelPlugin));
        Assert.Single(SampleMarkdownRenderer.StatusPanelFeaturePack.Plugins);
        Assert.Single(SampleMarkdownRenderer.StatusPanelFeaturePack.ReaderDocumentTransforms);
        Assert.Same(SampleMarkdownRenderer.StatusBadgeReaderDocumentTransform, SampleMarkdownRenderer.StatusPanelFeaturePack.ReaderDocumentTransforms[0]);
        Assert.Single(SampleMarkdownRenderer.StatusPanelFeaturePack.HtmlDocumentTransforms);
        Assert.Same(SampleMarkdownRenderer.StatusPanelHtmlDocumentTransform, SampleMarkdownRenderer.StatusPanelFeaturePack.HtmlDocumentTransforms[0]);
        Assert.Empty(SampleMarkdownRenderer.StatusPanelFeaturePack.RendererDocumentTransforms);
        Assert.Single(SampleMarkdownRenderer.StatusPanelFeaturePack.HtmlElementBlockConverters);
        Assert.Same(SampleMarkdownRenderer.StatusPanelVendorHtmlConverter, SampleMarkdownRenderer.StatusPanelFeaturePack.HtmlElementBlockConverters[0]);
        Assert.Single(SampleMarkdownRenderer.StatusPanelFeaturePack.HtmlInlineElementConverters);
        Assert.Same(SampleMarkdownRenderer.StatusBadgeInlineConverter, SampleMarkdownRenderer.StatusPanelFeaturePack.HtmlInlineElementConverters[0]);
        Assert.Single(SampleMarkdownRenderer.StatusPanelFeaturePack.VisualElementRoundTripHints);
        Assert.Contains(options.ReaderOptions.DocumentTransforms, transform => ReferenceEquals(transform, SampleMarkdownRenderer.StatusBadgeReaderDocumentTransform));
    }

    [Fact]
    public void MarkdownReaderOptions_Can_Apply_Renderer_Plugin_Reader_Contract_Idempotently() {
        var options = MarkdownReaderOptions.CreatePortableProfile();

        options.ApplyPlugin(SampleMarkdownRenderer.StatusPanelPlugin);
        options.ApplyPlugin(SampleMarkdownRenderer.StatusPanelPlugin);

        Assert.True(options.HasPlugin(SampleMarkdownRenderer.StatusPanelPlugin));
        Assert.Same(
            SampleMarkdownRenderer.StatusBadgeReaderDocumentTransform,
            Assert.Single(options.DocumentTransforms, transform => ReferenceEquals(transform, SampleMarkdownRenderer.StatusBadgeReaderDocumentTransform)));
    }

    [Fact]
    public void MarkdownReaderOptions_Can_Apply_Renderer_FeaturePack_Reader_Contract_Idempotently() {
        var options = MarkdownReaderOptions.CreatePortableProfile();

        options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);
        options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);

        Assert.True(options.HasFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack));
        Assert.True(options.HasPlugin(SampleMarkdownRenderer.StatusPanelPlugin));
        Assert.Same(
            SampleMarkdownRenderer.StatusBadgeReaderDocumentTransform,
            Assert.Single(options.DocumentTransforms, transform => ReferenceEquals(transform, SampleMarkdownRenderer.StatusBadgeReaderDocumentTransform)));
    }

    [Fact]
    public void SampleMarkdownRenderer_StatusBadgeReaderTransform_Upgrades_Source_Tokens_To_Typed_Inline_Ast() {
        var options = MarkdownReaderOptions.CreatePortableProfile();
        options.ApplyFeaturePack(SampleMarkdownRenderer.StatusPanelFeaturePack);

        MarkdownDoc document = MarkdownReader.Parse("System {{status:Healthy}} now", options);

        var paragraph = Assert.IsType<ParagraphBlock>(Assert.Single(document.Blocks));
        var highlight = Assert.IsType<HighlightInline>(Assert.Single(paragraph.Inlines.Nodes.OfType<HighlightInline>()));
        Assert.Equal("Healthy", highlight.Text);
        Assert.Equal("System ==Healthy== now", document.ToMarkdown().Trim());
    }

    [Fact]
    public void MarkdownRendererFeaturePack_Can_Compose_Plugins_With_Fence_Option_Schemas() {
        var schema = new MarkdownFenceOptionSchema(
            "vendor.visual-options",
            "Vendor Visual Options",
            new[] { "vendor-chart" },
            new[] {
                MarkdownFenceOptionDefinition.Boolean("pinned")
            });

        var plugin = new MarkdownRendererPlugin(
            "Vendor Visuals",
            new Func<MarkdownFencedCodeBlockRenderer>[] {
                () => new MarkdownFencedCodeBlockRenderer(
                    "Vendor chart",
                    new[] { "vendor-chart" },
                    (_, _) => "<div class=\"vendor-chart\"></div>")
            },
            new[] { schema });

        var featurePack = new MarkdownRendererFeaturePack(
            "vendor.visual-pack",
            "Vendor Visual Pack",
            new[] { plugin });

        var options = new MarkdownRendererOptions();
        options.ApplyFeaturePack(featurePack);

        Assert.True(options.HasFeaturePack(featurePack));
        Assert.True(options.HasPlugin(plugin));
        Assert.True(options.HasFenceOptionSchema(schema));
        Assert.True(options.TryParseFenceOptions("vendor-chart", MarkdownCodeFenceInfo.Parse("vendor-chart pinned"), out var parsed));
        Assert.True(parsed.TryGetBoolean("pinned", out var pinned));
        Assert.True(pinned);
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
    public void MarkdownRenderer_Custom_Renderers_Can_Read_Parsed_Fence_Metadata() {
        var md = """
```ix-note title="Release Note" pinned maxItems=3
hello
```
""";
        var opts = new MarkdownRendererOptions();
        opts.FencedCodeBlockRenderers.Add(new MarkdownFencedCodeBlockRenderer(
            "IX note metadata",
            new[] { "ix-note" },
            (match, _) => {
                var isPinned = match.FenceInfo.TryGetBooleanAttribute("pinned", out var pinned) && pinned;
                var maxItems = match.FenceInfo.TryGetInt32Attribute("maxItems", out var parsedMaxItems)
                    ? parsedMaxItems.ToString(System.Globalization.CultureInfo.InvariantCulture)
                    : string.Empty;
                return $"<aside class=\"ix-note\" data-lang=\"{System.Net.WebUtility.HtmlEncode(match.Language)}\" data-title=\"{System.Net.WebUtility.HtmlEncode(match.FenceInfo.Title)}\" data-pinned=\"{System.Net.WebUtility.HtmlEncode(isPinned.ToString().ToLowerInvariant())}\" data-max-items=\"{System.Net.WebUtility.HtmlEncode(maxItems)}\">{System.Net.WebUtility.HtmlEncode(match.RawContent)}</aside>";
            }));

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("class=\"ix-note\"", html, StringComparison.Ordinal);
        Assert.Contains("data-lang=\"ix-note\"", html, StringComparison.Ordinal);
        Assert.Contains("data-title=\"Release Note\"", html, StringComparison.Ordinal);
        Assert.Contains("data-pinned=\"true\"", html, StringComparison.Ordinal);
        Assert.Contains("data-max-items=\"3\"", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Custom_Renderers_Can_Read_Brace_Style_Fence_Metadata() {
        var md = """
```ix-note {#release-note .callout .pinned title="Release Note"}
hello
```
""";
        var opts = new MarkdownRendererOptions();
        opts.FencedCodeBlockRenderers.Add(new MarkdownFencedCodeBlockRenderer(
            "IX note metadata classes",
            new[] { "ix-note" },
            (match, _) => $"<aside class=\"ix-note\" data-id=\"{System.Net.WebUtility.HtmlEncode(match.FenceInfo.ElementId)}\" data-classes=\"{System.Net.WebUtility.HtmlEncode(string.Join(" ", match.FenceInfo.Classes))}\" data-title=\"{System.Net.WebUtility.HtmlEncode(match.FenceInfo.Title)}\">{System.Net.WebUtility.HtmlEncode(match.RawContent)}</aside>"));

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, opts);

        Assert.Contains("data-id=\"release-note\"", html, StringComparison.Ordinal);
        Assert.Contains("data-classes=\"callout pinned\"", html, StringComparison.Ordinal);
        Assert.Contains("data-title=\"Release Note\"", html, StringComparison.Ordinal);
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
    public void MarkdownRenderer_Dataview_Falls_Back_To_Fence_Title_Metadata_When_Json_Title_Is_Missing() {
        var md = """
```dataview title="Fallback Caption"
{"rows":[["Server","Fails"],["AD0","0"],["AD1","1"]]}
```
""";

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, MarkdownRendererPresets.CreateStrictMinimal());

        Assert.Contains("data-omd-dataview-title=\"Fallback Caption\"", html, StringComparison.Ordinal);
        Assert.Contains("<caption>Fallback Caption</caption>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_Dataview_Honors_Brace_Style_Fence_Id_And_Classes() {
        var md = """
```dataview {#replication-summary .wide .compact title="Replication Summary"}
{"rows":[["Server","Fails"],["AD0","0"],["AD1","1"]]}
```
""";

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml(md, MarkdownRendererPresets.CreateStrictMinimal());

        Assert.Contains("id=\"replication-summary\"", html, StringComparison.Ordinal);
        Assert.Contains("class=\"omd-visual omd-dataview wide compact\"", html, StringComparison.Ordinal);
        Assert.Contains("data-omd-dataview-title=\"Replication Summary\"", html, StringComparison.Ordinal);
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
    public void MarkdownRenderer_Can_Apply_Ast_Document_Transforms_Before_Html_Rendering() {
        var opts = new MarkdownRendererOptions();
        opts.DocumentTransforms.Add(new RendererAppendParagraphTransform("tail"));

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("hello", opts);

        Assert.Contains("<p>hello</p>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("<p>tail</p>", htmlOut, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRenderer_ParseDocument_Returns_Renderer_Transformed_Ast() {
        var opts = new MarkdownRendererOptions();
        opts.DocumentTransforms.Add(new RendererAppendParagraphTransform("tail"));

        var document = MarkdownRenderer.MarkdownRenderer.ParseDocument("hello", opts);

        Assert.Equal(2, document.Blocks.Count);
        Assert.Equal("hello", Assert.IsType<ParagraphBlock>(document.Blocks[0]).Inlines.RenderMarkdown());
        Assert.Equal("tail", Assert.IsType<ParagraphBlock>(document.Blocks[1]).Inlines.RenderMarkdown());
    }

    [Fact]
    public void MarkdownRenderer_ParseDocument_Can_Report_Transform_Diagnostics() {
        var opts = new MarkdownRendererOptions();
        opts.DocumentTransforms.Add(new RendererAppendParagraphTransform("tail"));
        var diagnostics = new List<MarkdownDocumentTransformDiagnostic>();

        var document = MarkdownRenderer.MarkdownRenderer.ParseDocument("hello", opts, diagnostics);

        Assert.Equal(2, document.Blocks.Count);
        var diagnostic = Assert.Single(diagnostics, diagnostic =>
            diagnostic.Source == MarkdownDocumentTransformSource.MarkdownRenderer
            && diagnostic.TransformName.Contains(nameof(RendererAppendParagraphTransform), StringComparison.Ordinal));
        Assert.Equal(1, diagnostic.BlockCountBefore);
        Assert.Equal(2, diagnostic.BlockCountAfter);
        Assert.False(diagnostic.ReplacedDocument);
        Assert.Equal(1, diagnostic.ChangedBlockStartBefore);
        Assert.Equal(0, diagnostic.ChangedBlockCountBefore);
        Assert.Equal(1, diagnostic.ChangedBlockStartAfter);
        Assert.Equal(1, diagnostic.ChangedBlockCountAfter);
        Assert.Null(diagnostic.AffectedSourceSpan);
    }

    [Fact]
    public void MarkdownRenderer_ParseDocument_Can_Report_PreProcessor_And_Transform_Diagnostics() {
        var opts = new MarkdownRendererOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };
        opts.DocumentTransforms.Add(new RendererAppendParagraphTransform("tail"));
        opts.MarkdownPreProcessors.Add((markdown, _) =>
            markdown.Replace("```mermaid\nflowchart LR", "```mermaid\ngraph TD"));
        var transformDiagnostics = new List<MarkdownDocumentTransformDiagnostic>();
        var preProcessorDiagnostics = new List<MarkdownRendererPreProcessorDiagnostic>();

        var document = MarkdownRenderer.MarkdownRenderer.ParseDocument(
            "```mermaidflowchart LR A-->B\n```",
            opts,
            transformDiagnostics,
            preProcessorDiagnostics);

        Assert.Equal(2, document.Blocks.Count);
        Assert.Equal(2, preProcessorDiagnostics.Count);
        Assert.Equal(MarkdownRendererPreProcessorStage.InputNormalization, preProcessorDiagnostics[0].Stage);
        Assert.Equal(MarkdownRendererPreProcessorStage.CustomPreProcessor, preProcessorDiagnostics[1].Stage);
        var diagnostic = Assert.Single(transformDiagnostics, diagnostic =>
            diagnostic.Source == MarkdownDocumentTransformSource.MarkdownRenderer
            && diagnostic.TransformName.Contains(nameof(RendererAppendParagraphTransform), StringComparison.Ordinal));
        Assert.Null(diagnostic.AffectedSourceSpan);
    }

    [Fact]
    public void MarkdownRenderer_ParseDocumentResult_Returns_SyntaxTree_And_Both_Diagnostic_Streams() {
        var opts = new MarkdownRendererOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };
        opts.DocumentTransforms.Add(new RendererAppendParagraphTransform("tail"));
        opts.MarkdownPreProcessors.Add((markdown, _) =>
            markdown.Replace("```mermaid\nflowchart LR", "```mermaid\ngraph TD"));

        var result = MarkdownRenderer.MarkdownRenderer.ParseDocumentResult("```mermaidflowchart LR A-->B\n```", opts);

        Assert.Equal(2, result.Document.Blocks.Count);
        Assert.Single(result.SyntaxTree.Children);
        Assert.Equal(2, result.PreProcessorDiagnostics.Count);
        Assert.Equal("```mermaid\ngraph TD A-->B\n```", result.PreprocessedMarkdown);
        Assert.Single(result.TransformDiagnostics, diagnostic =>
            diagnostic.Source == MarkdownDocumentTransformSource.MarkdownRenderer
            && diagnostic.TransformName.Contains(nameof(RendererAppendParagraphTransform), StringComparison.Ordinal));
    }

    [Fact]
    public void MarkdownRenderer_ParseDocumentResult_Provides_Final_SyntaxTree_And_Lookup_Helpers() {
        var opts = new MarkdownRendererOptions();
        opts.DocumentTransforms.Add(new RendererRewriteFirstParagraphTransform("hello renderer"));

        var result = MarkdownRenderer.MarkdownRenderer.ParseDocumentResult("hello", opts);

        Assert.Single(result.SyntaxTree.Children);
        Assert.Single(result.FinalSyntaxTree.Children);
        Assert.Equal("hello", result.FindDeepestNodeAtLine(1)!.Literal);
        Assert.Equal("hello renderer", result.FindDeepestFinalNodeAtLine(1)!.Literal);
        Assert.Equal("hello", result.FindDeepestNodeContainingSpan(new MarkdownSourceSpan(1, 1))!.Literal);
        Assert.Equal("hello renderer", result.FindDeepestFinalNodeContainingSpan(new MarkdownSourceSpan(1, 1))!.Literal);
        Assert.Equal(new[] { MarkdownSyntaxKind.Document, MarkdownSyntaxKind.Paragraph }, result.FindFinalNodePathAtLine(1).Select(node => node.Kind).ToArray());
        Assert.Equal("hello renderer", result.FindNearestFinalBlockOverlappingSpan(new MarkdownSourceSpan(1, 1))!.Literal);
    }

    [Fact]
    public void MarkdownRenderer_ParseDocumentResult_Provides_Position_Based_Syntax_Lookups() {
        var result = MarkdownRenderer.MarkdownRenderer.ParseDocumentResult("Use **bold** [docs](https://example.com) and `code`.", new MarkdownRendererOptions());

        Assert.Equal(MarkdownSyntaxKind.InlineText, result.FindDeepestNodeAtPosition(1, 8)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.InlineLinkTarget, result.FindDeepestNodeAtPosition(1, 30)!.Kind);
        Assert.Equal(MarkdownSyntaxKind.InlineCodeSpan, result.FindDeepestNodeAtPosition(1, 48)!.Kind);
        Assert.Equal(new[] { MarkdownSyntaxKind.Document, MarkdownSyntaxKind.Paragraph, MarkdownSyntaxKind.InlineLink, MarkdownSyntaxKind.InlineLinkTarget }, result.FindNodePathAtPosition(1, 30).Select(node => node.Kind).ToArray());
        Assert.Equal(MarkdownSyntaxKind.Paragraph, result.FindNearestBlockAtPosition(1, 48)!.Kind);
    }

    [Fact]
    public void MarkdownRenderer_ParseDocumentResult_Includes_Reader_And_Renderer_Transform_Diagnostics() {
        var opts = new MarkdownRendererOptions();
        opts.ReaderOptions.DocumentTransforms.Add(new ReaderAppendParagraphTransform("reader tail"));
        opts.DocumentTransforms.Add(new RendererAppendParagraphTransform("renderer tail"));

        var result = MarkdownRenderer.MarkdownRenderer.ParseDocumentResult("hello", opts);

        Assert.Equal(3, result.Document.Blocks.Count);
        Assert.True(result.TransformDiagnostics.Count >= 3);
        var readerDiagnostic = Assert.Single(result.TransformDiagnostics, diagnostic =>
            diagnostic.Source == MarkdownDocumentTransformSource.MarkdownReader
            && diagnostic.TransformName.Contains(nameof(ReaderAppendParagraphTransform), StringComparison.Ordinal));
        var rendererDiagnostic = Assert.Single(result.TransformDiagnostics, diagnostic =>
            diagnostic.Source == MarkdownDocumentTransformSource.MarkdownRenderer
            && diagnostic.TransformName.Contains(nameof(RendererAppendParagraphTransform), StringComparison.Ordinal));
        Assert.Null(readerDiagnostic.AffectedSourceSpan);
        Assert.Null(rendererDiagnostic.AffectedSourceSpan);
    }

    [Fact]
    public void MarkdownRenderer_ParseDocumentResult_Preserves_SourceSpans_For_RendererDiagnostics_After_ReaderBlockInsertions() {
        var opts = new MarkdownRendererOptions();
        opts.ReaderOptions.DocumentTransforms.Add(new ReaderAppendParagraphTransform("reader tail"));
        opts.DocumentTransforms.Add(new RendererRewriteFirstParagraphTransform("hello renderer"));

        var result = MarkdownRenderer.MarkdownRenderer.ParseDocumentResult("hello", opts);

        Assert.Equal(2, result.Document.Blocks.Count);
        var rendererDiagnostic = Assert.Single(result.TransformDiagnostics, diagnostic =>
            diagnostic.Source == MarkdownDocumentTransformSource.MarkdownRenderer
            && diagnostic.TransformName.Contains(nameof(RendererRewriteFirstParagraphTransform), StringComparison.Ordinal));
        Assert.Equal(new MarkdownSourceSpan(1, 1, 1, 5), rendererDiagnostic.AffectedSourceSpan);
    }

    [Fact]
    public void MarkdownRenderer_ParseDocumentResult_Preserves_SourceSpans_When_FrontMatter_Is_Present() {
        var opts = new MarkdownRendererOptions();
        opts.DocumentTransforms.Add(new RendererRewriteSecondParagraphTransform("second renderer"));

        var result = MarkdownRenderer.MarkdownRenderer.ParseDocumentResult("""
---
title: Sample
---

first

second
""", opts);

        var rendererDiagnostic = Assert.Single(result.TransformDiagnostics, diagnostic =>
            diagnostic.Source == MarkdownDocumentTransformSource.MarkdownRenderer
            && diagnostic.TransformName.Contains(nameof(RendererRewriteSecondParagraphTransform), StringComparison.Ordinal));
        Assert.Equal(new MarkdownSourceSpan(7, 1, 7, 6), rendererDiagnostic.AffectedSourceSpan);
    }

    [Fact]
    public void MarkdownRenderer_RenderBodyHtml_Does_Not_Mutate_Caller_HtmlOptions_BaseUri() {
        var opts = new MarkdownRendererOptions {
            BaseHref = "https://example.com/docs/"
        };

        Assert.Null(opts.HtmlOptions.BaseUri);

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("[x](page.html)", opts);

        Assert.Contains("<base href=\"https://example.com/docs/\">", html);
        Assert.Null(opts.HtmlOptions.BaseUri);
    }

    [Fact]
    public void MarkdownRendererPlugin_Can_Carry_Renderer_Document_Transforms_Idempotently() {
        var rendererTransform = new RendererAppendParagraphTransform("plugin tail");
        var plugin = new MarkdownRendererPlugin(
            "Vendor Renderer Visuals",
            new Func<MarkdownFencedCodeBlockRenderer>[] {
                () => new MarkdownFencedCodeBlockRenderer(
                    "Vendor chart",
                    new[] { "vendor-chart" },
                    (_, _) => "<div class=\"vendor-chart\"></div>")
            },
            rendererDocumentTransforms: new[] { rendererTransform });

        var options = new MarkdownRendererOptions();
        options.ApplyPlugin(plugin);
        options.ApplyPlugin(plugin);

        Assert.Same(rendererTransform, Assert.Single(options.DocumentTransforms, transform => ReferenceEquals(transform, rendererTransform)));

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("hello", options);
        Assert.Contains("<p>plugin tail</p>", html, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRendererFeaturePack_Composes_Renderer_Document_Transforms_From_Plugins() {
        var rendererTransform = new RendererAppendParagraphTransform("feature tail");
        var plugin = new MarkdownRendererPlugin(
            "Vendor Renderer Visuals",
            new Func<MarkdownFencedCodeBlockRenderer>[] {
                () => new MarkdownFencedCodeBlockRenderer(
                    "Vendor chart",
                    new[] { "vendor-chart" },
                    (_, _) => "<div class=\"vendor-chart\"></div>")
            },
            rendererDocumentTransforms: new[] { rendererTransform });
        var featurePack = new MarkdownRendererFeaturePack(
            "vendor.renderer-pack",
            "Vendor Renderer Pack",
            new[] { plugin });

        var options = new MarkdownRendererOptions();
        options.ApplyFeaturePack(featurePack);
        options.ApplyFeaturePack(featurePack);

        Assert.Same(rendererTransform, Assert.Single(featurePack.RendererDocumentTransforms, transform => ReferenceEquals(transform, rendererTransform)));
        Assert.Same(rendererTransform, Assert.Single(options.DocumentTransforms, transform => ReferenceEquals(transform, rendererTransform)));

        var html = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("hello", options);
        Assert.Contains("<p>feature tail</p>", html, StringComparison.Ordinal);
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
    public void MarkdownRenderer_Normalizes_List_Strong_Artifacts_When_Enabled() {
        var opts = new MarkdownRendererOptions {
            NormalizeLooseStrongDelimiters = true,
            NormalizeDanglingTrailingStrongListClosers = true,
            NormalizeMetricValueStrongRuns = true
        };

        var htmlOut = MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("""
- Overall health ****healthy****
- Overall health ✅ Healthy****
- Overall health ******healthy**
- Overall health **✅****Healthy**
- LDAP/LDAPS across all DCs **healthy on FQDN endpoints for all 5 servers*
""", opts);

        Assert.Contains("<strong>healthy</strong>", htmlOut, StringComparison.Ordinal);
        Assert.Contains("Healthy", htmlOut, StringComparison.Ordinal);
        Assert.Contains("healthy on FQDN endpoints for all 5 servers", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("Healthy****", htmlOut, StringComparison.Ordinal);
        Assert.DoesNotContain("******healthy**", htmlOut, StringComparison.Ordinal);
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
    public void MarkdownRenderer_Respects_Reader_MaxInputCharacters() {
        var opts = new MarkdownRendererOptions();
        opts.ReaderOptions.MaxInputCharacters = 8;

        var ex = Assert.Throws<ArgumentOutOfRangeException>(() =>
            MarkdownRenderer.MarkdownRenderer.RenderBodyHtml("123456789", opts));

        Assert.Contains("MaxInputCharacters", ex.Message, StringComparison.Ordinal);
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
    public void MarkdownRendererPreProcessorPipeline_Mirrors_Renderer_PreParse_Normalization_Order() {
        var opts = new MarkdownRendererOptions {
            NormalizeCompactFenceBodyBoundaries = true
        };
        opts.MarkdownPreProcessors.Add((markdown, _) =>
            markdown.Replace("```mermaid\nflowchart LR", "```mermaid\ngraph TD"));

        var processed = MarkdownRendererPreProcessorPipeline.Apply("```mermaidflowchart LR A-->B\n```", opts);

        Assert.Contains("```mermaid\ngraph TD A-->B", processed, StringComparison.Ordinal);
        Assert.DoesNotContain("```mermaidflowchart LR", processed, StringComparison.Ordinal);
    }

    [Fact]
    public void MarkdownRendererPreProcessorPipeline_Can_Report_Diagnostics() {
        var opts = new MarkdownRendererOptions {
            NormalizeEscapedNewlines = true,
            NormalizeCompactFenceBodyBoundaries = true
        };
        opts.MarkdownPreProcessors.Add((markdown, _) => markdown.Replace("graph TD", "flowchart LR"));
        var diagnostics = new List<MarkdownRendererPreProcessorDiagnostic>();

        var processed = MarkdownRendererPreProcessorPipeline.Apply(
            "```mermaidgraph TD A-->B\\n```",
            opts,
            diagnostics);

        Assert.Equal("```mermaid\nflowchart LR A-->B\n```", processed);
        Assert.Equal(3, diagnostics.Count);
        Assert.Equal(MarkdownRendererPreProcessorStage.EscapedNewlineNormalization, diagnostics[0].Stage);
        Assert.Equal(MarkdownRendererPreProcessorStage.InputNormalization, diagnostics[1].Stage);
        Assert.Equal(MarkdownRendererPreProcessorStage.CustomPreProcessor, diagnostics[2].Stage);
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

    private sealed class RendererAppendParagraphTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            Assert.Equal(MarkdownDocumentTransformSource.MarkdownRenderer, context.Source);
            Assert.NotNull(context.ReaderOptions);
            Assert.IsType<MarkdownRendererOptions>(context.SourceOptions);

            document.Add(new ParagraphBlock(new InlineSequence().Text(text)));
            return document;
        }
    }

    private sealed class ReaderAppendParagraphTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            Assert.Equal(MarkdownDocumentTransformSource.MarkdownReader, context.Source);
            Assert.NotNull(context.ReaderOptions);

            document.Add(new ParagraphBlock(new InlineSequence().Text(text)));
            return document;
        }
    }

    private sealed class RendererRewriteFirstParagraphTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            Assert.Equal(MarkdownDocumentTransformSource.MarkdownRenderer, context.Source);
            Assert.NotNull(context.ReaderOptions);
            Assert.IsType<MarkdownRendererOptions>(context.SourceOptions);

            var rewritten = MarkdownDoc.Create();
            rewritten.Add(new ParagraphBlock(new InlineSequence().Text(text)));
            for (var i = 1; i < document.Blocks.Count; i++) {
                rewritten.Add(document.Blocks[i]);
            }

            return rewritten;
        }
    }

    private sealed class RendererRewriteSecondParagraphTransform(string text) : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            Assert.Equal(MarkdownDocumentTransformSource.MarkdownRenderer, context.Source);
            Assert.NotNull(context.ReaderOptions);
            Assert.IsType<MarkdownRendererOptions>(context.SourceOptions);

            var rewritten = MarkdownDoc.Create();
            if (document.DocumentHeader != null) {
                rewritten.Add(document.DocumentHeader);
            }

            for (var i = 0; i < document.Blocks.Count; i++) {
                if (i == 1) {
                    rewritten.Add(new ParagraphBlock(new InlineSequence().Text(text)));
                } else {
                    rewritten.Add(document.Blocks[i]);
                }
            }

            return rewritten;
        }
    }

}

