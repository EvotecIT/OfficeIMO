using System.Security.Cryptography;
using System.Text.RegularExpressions;
using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

/// <summary>
/// Renders Markdown to HTML suitable for WebView2/browser hosts, and provides a reusable shell page
/// + an incremental update mechanism.
/// </summary>
public static class MarkdownRenderer {
    private static readonly Regex MermaidPreCodeBlockRegex = new Regex(
        "(<pre[^>]*>)\\s*<code\\s+class=\"language-mermaid\"[^>]*>([\\s\\S]*?)</code>\\s*</pre>",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex ChartPreCodeBlockRegex = new Regex(
        "(<pre[^>]*>)\\s*<code\\s+class=\"language-chart\"[^>]*>([\\s\\S]*?)</code>\\s*</pre>",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    private static readonly Regex MathPreCodeBlockRegex = new Regex(
        "(<pre[^>]*>)\\s*<code\\s+class=\"language-(math|latex)\"[^>]*>([\\s\\S]*?)</code>\\s*</pre>",
        RegexOptions.IgnoreCase | RegexOptions.Compiled);

    /// <summary>
    /// Parses Markdown using OfficeIMO.Markdown and returns an HTML fragment (typically an &lt;article class="markdown-body"&gt; wrapper).
    /// When Mermaid is enabled, Mermaid code blocks are annotated with hashes for incremental rendering.
    /// </summary>
    public static string RenderBodyHtml(string markdown, MarkdownRendererOptions? options = null) {
        options ??= new MarkdownRendererOptions();
        var readerOptions = options.ReaderOptions ?? new MarkdownReaderOptions();
        var htmlOptions = options.HtmlOptions ?? new HtmlOptions { Kind = HtmlKind.Fragment };

        if (options.NormalizeEscapedNewlines && !string.IsNullOrEmpty(markdown)) {
            markdown = markdown.Replace("\\r\\n", "\n").Replace("\\n", "\n");
        }

        if (!string.IsNullOrWhiteSpace(options.BaseHref) && htmlOptions.BaseUri == null) {
            // Best-effort: use BaseHref for origin restrictions (if enabled). If parsing fails or BaseHref isn't absolute,
            // keep BaseUri null and origin restriction will effectively be disabled.
            if (Uri.TryCreate(options.BaseHref!.Trim(), UriKind.Absolute, out var baseUri)) {
                htmlOptions.BaseUri = baseUri;
            }
        }

        var doc = MarkdownReader.Parse(markdown ?? string.Empty, readerOptions);
        string html = doc.ToHtmlFragment(htmlOptions);

        if (options.Mermaid?.Enabled == true) {
            html = ConvertMermaidCodeBlocks(html, enableHashCaching: options.Mermaid.EnableHashCaching);
        }

        if (options.Chart?.Enabled == true) {
            html = ConvertChartCodeBlocks(html);
        }

        if (options.Math?.Enabled == true && options.Math.EnableFencedMathBlocks) {
            html = ConvertMathCodeBlocks(html, options.Math);
        }

        if (!string.IsNullOrWhiteSpace(options.BaseHref)) {
            // Put <base> into the update payload. The incremental updater moves it into <head>.
            var baseHref = System.Net.WebUtility.HtmlEncode(options.BaseHref!.Trim());
            html = $"<base href=\"{baseHref}\">" + html;
        }

        return html;
    }

    /// <summary>
    /// Builds a self-contained HTML document that preloads CSS and scripts once (Prism/Mermaid),
    /// and exposes a global <c>updateContent(newBodyHtml)</c> function for incremental updates.
    /// </summary>
    public static string BuildShellHtml(string? title = null, MarkdownRendererOptions? options = null) {
        options ??= new MarkdownRendererOptions();
        var htmlOptions = options.HtmlOptions ?? new HtmlOptions { Kind = HtmlKind.Fragment };

        // Build head assets (CSS + optional Prism assets) from OfficeIMO.Markdown.
        // This intentionally uses an empty doc; content is pushed later via updateContent(...).
        var empty = MarkdownDoc.Create();
        var parts = empty.ToHtmlParts(htmlOptions);

        var sb = new StringBuilder(16 * 1024);
        sb.Append("<!DOCTYPE html><html lang=\"en\"><head><meta charset=\"utf-8\"><meta name=\"viewport\" content=\"width=device-width,initial-scale=1\">");
        if (!string.IsNullOrWhiteSpace(options.ContentSecurityPolicy)) {
            sb.Append("<meta http-equiv=\"Content-Security-Policy\" content=\"")
              .Append(System.Net.WebUtility.HtmlEncode(options.ContentSecurityPolicy!.Trim()))
              .Append("\">");
        }
        sb.Append("<title>").Append(System.Net.WebUtility.HtmlEncode(title ?? "Markdown")).Append("</title>");
        if (!string.IsNullOrEmpty(parts.Css)) sb.Append("<style>\n").Append(parts.Css).Append("\n</style>");
        if (!string.IsNullOrEmpty(parts.Head)) sb.Append(parts.Head);

        if (options.Math?.Enabled == true) {
            sb.Append(BuildMathBootstrap(options.Math));
        }

        if (options.Mermaid?.Enabled == true) {
            sb.Append(BuildMermaidBootstrap(options.Mermaid));
        }

        if (options.Chart?.Enabled == true) {
            sb.Append(BuildChartBootstrap(options.Chart));
        }

        sb.Append("</head><body>");
        sb.Append("<div id=\"omdRoot\"></div>");
        sb.Append("<script>\n").Append(BuildIncrementalUpdateScript(options)).Append("\n</script>");
        sb.Append("</body></html>");
        return sb.ToString();
    }

    /// <summary>
    /// Returns a JavaScript snippet that calls <c>updateContent(...)</c> with a properly escaped string literal.
    /// </summary>
    public static string BuildUpdateScript(string bodyHtml) {
        return "updateContent(" + JavaScriptString.SingleQuoted(bodyHtml ?? string.Empty) + ");";
    }

    private static string ConvertMermaidCodeBlocks(string html, bool enableHashCaching) {
        if (string.IsNullOrEmpty(html)) return html;
        // Mermaid expects elements with class="mermaid" containing the diagram text. Convert fenced blocks
        // rendered as <pre><code class="language-mermaid">..</code></pre> into <pre class="mermaid">..</pre>.
        return MermaidPreCodeBlockRegex.Replace(html, m => {
            var content = m.Groups[2].Value;
            string hashAttr = string.Empty;
            if (enableHashCaching) {
                string hash = ComputeShortHash(content);
                hashAttr = $" data-mermaid-hash=\"{hash}\"";
            }
            return $"<pre class=\"mermaid\"{hashAttr}>{content}</pre>";
        });
    }

    private static string ConvertChartCodeBlocks(string html) {
        if (string.IsNullOrEmpty(html)) return html;
        // Charts are authored as fenced code blocks named `chart` with JSON config. Convert
        // <pre><code class="language-chart">..</code></pre> into a <canvas> annotated with base64 config.
        return ChartPreCodeBlockRegex.Replace(html, m => {
            var encoded = m.Groups[2].Value ?? string.Empty;
            var rawJson = System.Net.WebUtility.HtmlDecode(encoded);
            var b64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(rawJson ?? string.Empty));
            var hash = ComputeShortHash(encoded);
            return $"<canvas class=\"omd-chart\" data-chart-hash=\"{hash}\" data-chart-config-b64=\"{System.Net.WebUtility.HtmlEncode(b64)}\"></canvas>";
        });
    }

    private static string ConvertMathCodeBlocks(string html, MathOptions mathOptions) {
        if (string.IsNullOrEmpty(html)) return html;
        // Convert fenced ```math/```latex blocks rendered as code fences into display-math text nodes.
        // KaTeX auto-render runs on the updated DOM and will render the $$...$$ delimiters.
        return MathPreCodeBlockRegex.Replace(html, m => {
            var lang = (m.Groups[2].Value ?? string.Empty).Trim();
            if (!IsMathFenceLanguageAllowed(lang, mathOptions)) return m.Value;

            var encoded = m.Groups[3].Value ?? string.Empty;
            var raw = System.Net.WebUtility.HtmlDecode(encoded) ?? string.Empty;

            // Re-encode to keep content safe as text. Preserve newlines for nicer display rendering.
            var safe = System.Net.WebUtility.HtmlEncode(raw);
            return "<div class=\"omd-math\">$$\n" + safe + "\n$$</div>";
        });
    }

    private static bool IsMathFenceLanguageAllowed(string lang, MathOptions mathOptions) {
        if (string.IsNullOrWhiteSpace(lang)) return false;
        if (mathOptions == null) return false;
        var allowed = mathOptions.FencedMathLanguages;
        if (allowed == null || allowed.Length == 0) return true; // treat as enabled for defaults

        for (int i = 0; i < allowed.Length; i++) {
            var a = (allowed[i] ?? string.Empty).Trim();
            if (a.Length == 0) continue;
            if (string.Equals(a, lang, StringComparison.OrdinalIgnoreCase)) return true;
        }
        return false;
    }

    private static string ComputeShortHash(string input) {
        var data = Encoding.UTF8.GetBytes(input ?? string.Empty);
        byte[] hash;
#if NET8_0_OR_GREATER
        hash = SHA256.HashData(data);
#else
        using (var sha = SHA256.Create()) {
            hash = sha.ComputeHash(data);
        }
#endif
        // Use first 8 bytes as hex = 16 chars, plenty for DOM-diff keys.
        return ToHex(hash, 8);
    }

    private static string ToHex(byte[] bytes, int take) {
        if (bytes == null || bytes.Length == 0) return string.Empty;
        int len = Math.Min(take, bytes.Length);
        var sb = new StringBuilder(len * 2);
        for (int i = 0; i < len; i++) {
            sb.Append(bytes[i].ToString("x2"));
        }
        return sb.ToString();
    }

    private static string BuildMermaidBootstrap(MermaidOptions o) {
        // Use ESM bootstrap for Mermaid.
        string url = System.Net.WebUtility.HtmlEncode((o?.EsmModuleUrl ?? string.Empty).Trim());
        string light = System.Net.WebUtility.HtmlEncode((o?.LightTheme ?? "default").Trim());
        string dark = System.Net.WebUtility.HtmlEncode((o?.DarkTheme ?? "dark").Trim());
        if (string.IsNullOrEmpty(url)) return string.Empty;
        return $@"
<script type=""module"">
import mermaid from '{url}';
window.mermaid = mermaid;
mermaid.initialize({{ startOnLoad: false, theme: window.matchMedia('(prefers-color-scheme: dark)').matches ? '{dark}' : '{light}' }});
</script>";
    }

    private static string BuildChartBootstrap(ChartOptions o) {
        string url = System.Net.WebUtility.HtmlEncode((o?.ScriptUrl ?? string.Empty).Trim());
        if (string.IsNullOrEmpty(url)) return string.Empty;
        return $"\n<script src=\"{url}\"></script>\n";
    }

    private static string BuildMathBootstrap(MathOptions o) {
        string css = System.Net.WebUtility.HtmlEncode((o?.CssUrl ?? string.Empty).Trim());
        string js = System.Net.WebUtility.HtmlEncode((o?.ScriptUrl ?? string.Empty).Trim());
        string ar = System.Net.WebUtility.HtmlEncode((o?.AutoRenderScriptUrl ?? string.Empty).Trim());
        if (string.IsNullOrEmpty(css) || string.IsNullOrEmpty(js) || string.IsNullOrEmpty(ar)) return string.Empty;

        // KaTeX should be ready before we render content via updateContent(...). Use defer so it doesn't block HTML parse,
        // and call renderMathInElement from updateContent after DOM updates.
        return $"\n<link rel=\"stylesheet\" href=\"{css}\">\n<script defer src=\"{js}\"></script>\n<script defer src=\"{ar}\"></script>\n";
    }

    private static string BuildIncrementalUpdateScript(MarkdownRendererOptions options) {
        bool mermaid = options.Mermaid?.Enabled == true;
        bool chart = options.Chart?.Enabled == true;
        var mathOptions = options.Math;

        // Notes:
        // - We keep <base> in <head> so relative links/images resolve.
        // - We preserve already-rendered Mermaid SVGs by comparing data-mermaid-hash attributes.
        // - We re-run Prism highlighting after updates (if Prism is present).
        var sb = new StringBuilder(8 * 1024);
        sb.Append("""
async function updateContent(newBodyHtml) {
  const root = document.getElementById('omdRoot') || document.body;
  // Extract <base href="..."> from payload and place it in <head>.
  try {
    const baseMatch = newBodyHtml.match(/<base\s+href="([^"]*)"[^>]*>/i);
    if (baseMatch) {
      let baseEl = document.getElementById('omdBase');
      if (!baseEl) {
        baseEl = document.createElement('base');
        baseEl.id = 'omdBase';
        document.head.appendChild(baseEl);
      }
      baseEl.href = baseMatch[1];
      newBodyHtml = newBodyHtml.replace(baseMatch[0], '');
    } else {
      const baseEl = document.getElementById('omdBase');
      if (baseEl) baseEl.href = 'about:blank';
    }
  } catch(e) { /* best-effort */ }
""");

        if (chart) {
            sb.Append("""
  // Destroy existing Chart.js instances before replacing DOM to avoid leaks.
  try {
    if (window.Chart && typeof Chart.getChart === 'function') {
      root.querySelectorAll('canvas.omd-chart').forEach(c => {
        try { const inst = Chart.getChart(c); if (inst) inst.destroy(); } catch(e) { /* ignore */ }
      });
    }
  } catch(e) { /* ignore */ }
""");
        }

        if (mermaid) {
            sb.Append("""
  // Cache existing Mermaid SVGs keyed by data-mermaid-hash.
  const existingSvgs = new Map();
  root.querySelectorAll('[data-mermaid-hash]').forEach(el => {
    const hash = el.getAttribute('data-mermaid-hash');
    const svg = el.querySelector('svg') || (el.nextElementSibling && el.nextElementSibling.tagName === 'svg' ? el.nextElementSibling : null);
    if (hash && svg) existingSvgs.set(hash, svg.cloneNode(true));
  });
""");
        }

        sb.Append("""
  // Replace rendered contents.
  root.innerHTML = newBodyHtml;
""");

        if (mermaid) {
            sb.Append("""
  // Restore cached Mermaid SVGs for unchanged diagrams.
  root.querySelectorAll('[data-mermaid-hash]').forEach(el => {
    const hash = el.getAttribute('data-mermaid-hash');
    if (existingSvgs.has(hash)) {
      const cachedSvg = existingSvgs.get(hash);
      el.innerHTML = '';
      el.appendChild(cachedSvg);
      el.setAttribute('data-mermaid-rendered', 'true');
    }
  });

  // Render only new/changed Mermaid blocks.
  const unrendered = root.querySelectorAll('[data-mermaid-hash]:not([data-mermaid-rendered])');
  if (unrendered.length > 0 && window.mermaid) {
    try { await window.mermaid.run({ nodes: unrendered }); } catch(e) { console.warn('Mermaid render error:', e); }
  }
  // Render plain Mermaid blocks (language-mermaid) when hashes are not present.
  const plainMermaid = root.querySelectorAll('pre > code.language-mermaid:not([data-mermaid-rendered]), .mermaid:not([data-mermaid-rendered]):not(svg)');
  if (plainMermaid.length > 0 && window.mermaid) {
    try { await window.mermaid.run({ nodes: plainMermaid }); } catch(e) { console.warn('Mermaid render error:', e); }
  }
""");
        }

        if (chart) {
            sb.Append("""
  // Chart.js rendering (optional).
  try {
    function b64ToUtf8(b64) {
      try {
        const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
        if (window.TextDecoder) return new TextDecoder('utf-8').decode(bytes);
        // Fallback for older engines.
        return decodeURIComponent(escape(String.fromCharCode.apply(null, Array.from(bytes))));
      } catch(e) { return ''; }
    }
    if (window.Chart) {
      root.querySelectorAll('canvas.omd-chart:not([data-chart-rendered])').forEach(c => {
        const b64 = c.getAttribute('data-chart-config-b64');
        if (!b64) return;
        const jsonText = b64ToUtf8(b64);
        if (!jsonText) return;
        let cfg = null;
        try { cfg = JSON.parse(jsonText); } catch(e) { console.warn('Chart config JSON parse error:', e); return; }
        try {
          const ctx = c.getContext && c.getContext('2d');
          if (!ctx) return;
          new Chart(ctx, cfg);
          c.setAttribute('data-chart-rendered', 'true');
        } catch(e) { console.warn('Chart render error:', e); }
      });
    }
  } catch(e) { /* ignore */ }
""");
        }

        sb.Append("""
  // Prism highlighting (optional).
  try {
    if (window.Prism) {
      if (typeof Prism.highlightAllUnder === 'function') Prism.highlightAllUnder(root);
      else if (typeof Prism.highlightAll === 'function') Prism.highlightAll();
    }
  } catch(e) { /* ignore */ }
""");

        if (mathOptions != null && mathOptions.Enabled) {
            sb.Append("""

  // KaTeX auto-render (optional).
  try {
    if (window.renderMathInElement) {
      const delimiters = [];
""");
            if (mathOptions.EnableDollarDisplay) sb.Append("      delimiters.push({ left: \"$$\", right: \"$$\", display: true });\n");
            if (mathOptions.EnableDollarInline) sb.Append("      delimiters.push({ left: \"$\", right: \"$\", display: false });\n");
            if (mathOptions.EnableBracketDisplay) sb.Append("      delimiters.push({ left: \"\\\\[\", right: \"\\\\]\", display: true });\n");
            if (mathOptions.EnableParenInline) sb.Append("      delimiters.push({ left: \"\\\\(\", right: \"\\\\)\", display: false });\n");
            sb.Append("""
      if (delimiters.length > 0) {
        window.renderMathInElement(root, {
          delimiters: delimiters,
          throwOnError: false,
          strict: 'ignore',
          ignoredTags: ['script', 'noscript', 'style', 'textarea', 'pre', 'code']
        });
      }
    }
  } catch(e) { /* ignore */ }
""");
        }

        sb.Append("}\n");
        return sb.ToString();
    }
}
