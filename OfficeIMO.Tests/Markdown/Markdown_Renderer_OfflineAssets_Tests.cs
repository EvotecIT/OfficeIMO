using System;
using System.IO;
using OfficeIMO.Markdown;
using OfficeIMO.MarkdownRenderer;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public class Markdown_Renderer_OfflineAssets_Tests {
    [Fact]
    public void MarkdownRenderer_Shell_Bundles_Mermaid_Chart_And_Math_Assets_When_Offline_And_Local_Paths_Are_Provided() {
        string tempDir = Path.Combine(Path.GetTempPath(), "OfficeIMO.MarkdownRenderer.Tests", Guid.NewGuid().ToString("n"));
        Directory.CreateDirectory(tempDir);

        string mermaidJs = Path.Combine(tempDir, "mermaid.min.js");
        string chartJs = Path.Combine(tempDir, "chart.umd.min.js");
        string katexCss = Path.Combine(tempDir, "katex.min.css");
        string katexJs = Path.Combine(tempDir, "katex.min.js");
        string autoRenderJs = Path.Combine(tempDir, "auto-render.min.js");

        try {
            // Keep these tiny; we're testing bundling mechanics (data: URLs), not library behavior.
            File.WriteAllText(mermaidJs, "window.mermaid = window.mermaid || { initialize: function(){} };", System.Text.Encoding.UTF8);
            File.WriteAllText(chartJs, "window.Chart = window.Chart || function(){};", System.Text.Encoding.UTF8);
            File.WriteAllText(katexCss, "/* katex css */ .katex{font:16px serif;}", System.Text.Encoding.UTF8);
            File.WriteAllText(katexJs, "window.katex = window.katex || {};", System.Text.Encoding.UTF8);
            File.WriteAllText(autoRenderJs, "window.renderMathInElement = window.renderMathInElement || function(){};", System.Text.Encoding.UTF8);

            var opts = new MarkdownRendererOptions {
                HtmlOptions = new HtmlOptions {
                    Kind = HtmlKind.Fragment,
                    Style = HtmlStyle.Plain,
                    CssDelivery = CssDelivery.None,
                    BodyClass = null,
                    AssetMode = AssetMode.Offline
                }
            };

            opts.Mermaid.Enabled = true;
            opts.Mermaid.ScriptUrl = mermaidJs;

            opts.Chart.Enabled = true;
            opts.Chart.ScriptUrl = chartJs;

            opts.Math.Enabled = true;
            opts.Math.CssUrl = katexCss;
            opts.Math.ScriptUrl = katexJs;
            opts.Math.AutoRenderScriptUrl = autoRenderJs;

            var shell = MarkdownRenderer.MarkdownRenderer.BuildShellHtml("Chat", opts);

            Assert.Contains("data:text/css;base64,", shell, StringComparison.OrdinalIgnoreCase);
            Assert.Contains("data:application/javascript;base64,", shell, StringComparison.OrdinalIgnoreCase);

            // Offline Mermaid should use the non-module path (no ESM import in shell).
            Assert.DoesNotContain("type=\"module\"", shell, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("mermaid.esm.min.mjs", shell, StringComparison.OrdinalIgnoreCase);

            // We expect the file paths not to leak into the HTML when bundling succeeds.
            Assert.DoesNotContain("mermaid.min.js", shell, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("chart.umd.min.js", shell, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("katex.min.css", shell, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("katex.min.js", shell, StringComparison.OrdinalIgnoreCase);
            Assert.DoesNotContain("auto-render.min.js", shell, StringComparison.OrdinalIgnoreCase);
        } finally {
            try { Directory.Delete(tempDir, recursive: true); } catch { /* best-effort */ }
        }
    }
}

