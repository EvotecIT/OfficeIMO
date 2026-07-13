using System.Security.Cryptography;
using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

public static partial class MarkdownRenderer {
    internal static string ComputeShortHash(string input) {
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

    private static string BuildMermaidBootstrap(MermaidOptions o, AssetMode assetMode, MarkdownExternalTextResolver? resolver) {
        // Mermaid bootstrap:
        // - Online: ESM import (default)
        // - Offline: non-module script (easier to bundle/host locally)
        string url = (o?.EsmModuleUrl ?? string.Empty).Trim();
        string scriptUrl = (o?.ScriptUrl ?? string.Empty).Trim();
        string light = (o?.LightTheme ?? "default").Trim();
        string dark = (o?.DarkTheme ?? "dark").Trim();
        if (string.IsNullOrEmpty(url) && string.IsNullOrEmpty(scriptUrl)) return string.Empty;

        // Prevent closing the <script> tag if a caller passes a hostile value.
        url = ReplaceScriptCloseSequence(url);
        scriptUrl = ReplaceScriptCloseSequence(scriptUrl);
        light = ReplaceScriptCloseSequence(light);
        dark = ReplaceScriptCloseSequence(dark);

        if (assetMode == AssetMode.Offline && !string.IsNullOrEmpty(scriptUrl)) {
            string src = BuildBundledScriptSrc(scriptUrl, mime: "application/javascript", resolver);
            if (string.IsNullOrEmpty(src)) src = scriptUrl;
            src = System.Net.WebUtility.HtmlEncode(src);
            return $@"
<script src=""{src}""></script>
<script>
// Initialize Mermaid once after load (non-module path).
(function(){{
  try {{
    if (!window.mermaid && typeof mermaid !== 'undefined') window.mermaid = mermaid;
    if (window.mermaid && typeof window.mermaid.initialize === 'function') {{
      window.mermaid.initialize({{ startOnLoad: false, theme: window.matchMedia('(prefers-color-scheme: dark)').matches ? {JavaScriptString.SingleQuoted(dark)} : {JavaScriptString.SingleQuoted(light)} }});
    }}
  }} catch(e) {{ }}
}})();
</script>";
        }

        // Default (online): ESM import.
        if (string.IsNullOrEmpty(url)) return string.Empty;
        return $@"
<script type=""module"">
import mermaid from {JavaScriptString.SingleQuoted(url)};
window.mermaid = mermaid;
mermaid.initialize({{ startOnLoad: false, theme: window.matchMedia('(prefers-color-scheme: dark)').matches ? {JavaScriptString.SingleQuoted(dark)} : {JavaScriptString.SingleQuoted(light)} }});
</script>";
    }

    private static string ReplaceScriptCloseSequence(string value) {
        if (string.IsNullOrEmpty(value)) return string.Empty;
        // Avoid embedding a literal "</script" inside script contents.
        return value.Replace("</", "<\\/");
    }

    private static string BuildChartBootstrap(ChartOptions o, AssetMode assetMode, MarkdownExternalTextResolver? resolver) {
        string url = (o?.ScriptUrl ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(url)) return string.Empty;

        string src = assetMode == AssetMode.Offline ? BuildBundledScriptSrc(url, mime: "application/javascript", resolver) : string.Empty;
        if (string.IsNullOrEmpty(src)) src = url;
        src = System.Net.WebUtility.HtmlEncode(src);
        return $"\n<script defer src=\"{src}\"></script>\n";
    }

    private static string BuildMathBootstrap(MathOptions o, AssetMode assetMode, MarkdownExternalTextResolver? resolver) {
        string css = (o?.CssUrl ?? string.Empty).Trim();
        string js = (o?.ScriptUrl ?? string.Empty).Trim();
        string ar = (o?.AutoRenderScriptUrl ?? string.Empty).Trim();
        if (string.IsNullOrEmpty(css) || string.IsNullOrEmpty(js) || string.IsNullOrEmpty(ar)) return string.Empty;

        // KaTeX should be ready before we render content via updateContent(...). Use defer so it doesn't block HTML parse,
        // and call renderMathInElement from updateContent after DOM updates.
        string cssHref = assetMode == AssetMode.Offline ? BuildBundledCssHref(css, resolver) : string.Empty;
        if (string.IsNullOrEmpty(cssHref)) cssHref = css;
        cssHref = System.Net.WebUtility.HtmlEncode(cssHref);

        string jsSrc = assetMode == AssetMode.Offline ? BuildBundledScriptSrc(js, mime: "application/javascript", resolver) : string.Empty;
        if (string.IsNullOrEmpty(jsSrc)) jsSrc = js;
        jsSrc = System.Net.WebUtility.HtmlEncode(jsSrc);

        string arSrc = assetMode == AssetMode.Offline ? BuildBundledScriptSrc(ar, mime: "application/javascript", resolver) : string.Empty;
        if (string.IsNullOrEmpty(arSrc)) arSrc = ar;
        arSrc = System.Net.WebUtility.HtmlEncode(arSrc);

        return $"\n<link rel=\"stylesheet\" href=\"{cssHref}\">\n<script defer src=\"{jsSrc}\"></script>\n<script defer src=\"{arSrc}\"></script>\n";
    }

    private static void AppendCustomShellHeadHtml(StringBuilder sb, MarkdownRendererOptions options, AssetMode assetMode) {
        var renderers = options.FencedCodeBlockRenderers;
        if (renderers == null || renderers.Count == 0) {
            return;
        }

        for (int i = 0; i < renderers.Count; i++) {
            var renderer = renderers[i];
            if (renderer?.BuildShellHeadHtml == null) {
                continue;
            }

            var fragment = renderer.BuildShellHeadHtml(options, assetMode);
            if (!string.IsNullOrWhiteSpace(fragment)) {
                sb.Append(fragment);
            }
        }
    }

    private static void AppendCustomUpdateScripts(StringBuilder sb, MarkdownRendererOptions options, bool beforeReplace) {
        var renderers = options.FencedCodeBlockRenderers;
        if (renderers == null || renderers.Count == 0) {
            return;
        }

        for (int i = 0; i < renderers.Count; i++) {
            var renderer = renderers[i];
            if (renderer == null) {
                continue;
            }

            var builder = beforeReplace
                ? renderer.BuildBeforeContentReplaceScript
                : renderer.BuildAfterContentReplaceScript;
            if (builder == null) {
                continue;
            }

            var fragment = builder(options);
            if (string.IsNullOrWhiteSpace(fragment)) {
                continue;
            }

            sb.Append('\n')
              .Append(ReplaceScriptCloseSequence(fragment ?? string.Empty))
              .Append('\n');
        }
    }

    internal static string BuildBundledScriptSrc(string hrefOrPath, string mime, MarkdownExternalTextResolver? resolver) {
        // Only used by shell building logic. This should never throw.
        try {
            var text = TryLoadTextAsset(hrefOrPath, resolver);
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var bytes = Encoding.UTF8.GetBytes(text);
            var b64 = Convert.ToBase64String(bytes);
            return $"data:{mime};base64,{b64}";
        } catch { return string.Empty; }
    }

    internal static string BuildBundledCssHref(string hrefOrPath, MarkdownExternalTextResolver? resolver) {
        try {
            var text = TryLoadTextAsset(hrefOrPath, resolver);
            if (string.IsNullOrEmpty(text)) return string.Empty;
            var bytes = Encoding.UTF8.GetBytes(text);
            var b64 = Convert.ToBase64String(bytes);
            return $"data:text/css;base64,{b64}";
        } catch { return string.Empty; }
    }

    private static string TryLoadTextAsset(string hrefOrPath, MarkdownExternalTextResolver? resolver) {
        try {
            if (string.IsNullOrWhiteSpace(hrefOrPath)) return string.Empty;
            string v = hrefOrPath.Trim();

            if (Uri.TryCreate(v, UriKind.Absolute, out var uri)) {
                if (uri.IsFile) {
                    string path = uri.LocalPath;
                    return TryReadAllTextBounded(path);
                }

                if (string.Equals(uri.Scheme, Uri.UriSchemeHttp, StringComparison.OrdinalIgnoreCase) ||
                    string.Equals(uri.Scheme, Uri.UriSchemeHttps, StringComparison.OrdinalIgnoreCase)) {
                    return TryResolveTextBounded(uri, resolver);
                }

                // Unknown scheme (e.g., custom WebView virtual hosts) cannot be resolved by this process.
                return string.Empty;
            }

            // Treat as local path (absolute or relative).
            return TryReadAllTextBounded(v);
        } catch {
            return string.Empty;
        }
    }

    private static string TryReadAllTextBounded(string path) {
        try {
            if (string.IsNullOrWhiteSpace(path)) return string.Empty;
            if (!System.IO.File.Exists(path)) return string.Empty;
            const long MaxBytes = 10_000_000; // 10MB guardrail
            var fi = new System.IO.FileInfo(path);
            if (fi.Length > MaxBytes) return string.Empty;
            return System.IO.File.ReadAllText(path, Encoding.UTF8);
        } catch {
            return string.Empty;
        }
    }

    private static string TryResolveTextBounded(Uri uri, MarkdownExternalTextResolver? resolver) {
        try {
            if (uri == null || resolver == null) return string.Empty;
            const int MaxCharacters = 10_000_000;
            string? text = resolver(uri);
            return text != null && text.Length <= MaxCharacters ? text : string.Empty;
        } catch {
            return string.Empty;
        }
    }
}
