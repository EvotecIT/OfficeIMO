using System.Text.Json;
using System.Text;
using OfficeIMO.Markdown;

namespace OfficeIMO.MarkdownRenderer;

internal static class MarkdownRendererBuiltInFencedCodeBlocks {
    public static void RegisterDefaults(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.FencedCodeBlockRenderers.Add(CreateChartRenderer());
        options.FencedCodeBlockRenderers.Add(CreateNetworkRenderer());
        options.FencedCodeBlockRenderers.Add(CreateDataViewRenderer());
    }

    private static MarkdownFencedCodeBlockRenderer CreateChartRenderer() {
        return new MarkdownFencedCodeBlockRenderer(
            "Built-in Chart.js",
            new[] { "chart", "ix-chart" },
            (match, options) => options.Chart?.Enabled == true
                ? BuildNativeVisualHtml("canvas", "omd-chart", "chart", match.Language, "data-chart-hash", "data-chart-config-b64", match.RawContent)
                : null);
    }

    private static MarkdownFencedCodeBlockRenderer CreateNetworkRenderer() {
        return new MarkdownFencedCodeBlockRenderer(
            "Built-in vis-network",
            new[] { "ix-network", "network", "visnetwork" },
            (match, options) => options.Network?.Enabled == true
                ? BuildNativeVisualHtml("div", "omd-network", "network", match.Language, "data-network-hash", "data-network-config-b64", match.RawContent)
                : null) {
            BuildShellHeadHtml = BuildNetworkShellHeadHtml,
            BuildBeforeContentReplaceScript = BuildNetworkBeforeReplaceScript,
            BuildAfterContentReplaceScript = BuildNetworkAfterReplaceScript
        };
    }

    private static MarkdownFencedCodeBlockRenderer CreateDataViewRenderer() {
        return new MarkdownFencedCodeBlockRenderer(
            "Built-in IX dataview",
            new[] { "ix-dataview" },
            (match, _) => TryBuildDataViewHtml(match.RawContent));
    }

    private static string BuildNativeVisualHtml(string elementName, string cssClass, string visualKind, string language, string hashAttribute, string configAttribute, string rawContent) {
        var raw = rawContent ?? string.Empty;
        var base64 = Convert.ToBase64String(Encoding.UTF8.GetBytes(raw));
        var hash = MarkdownRenderer.ComputeShortHash(raw);
        var encodedClass = System.Net.WebUtility.HtmlEncode("omd-visual " + (cssClass ?? string.Empty).Trim());
        var encodedKind = System.Net.WebUtility.HtmlEncode(visualKind ?? string.Empty);
        var encodedLanguage = System.Net.WebUtility.HtmlEncode(language ?? string.Empty);
        var encodedHash = System.Net.WebUtility.HtmlEncode(hash);
        var encodedBase64 = System.Net.WebUtility.HtmlEncode(base64);
        return $"<{elementName} class=\"{encodedClass}\" data-omd-visual-contract=\"v1\" data-omd-visual-kind=\"{encodedKind}\" data-omd-fence-language=\"{encodedLanguage}\" data-omd-visual-hash=\"{encodedHash}\" data-omd-config-format=\"json\" data-omd-config-encoding=\"base64-utf8\" data-omd-config-b64=\"{encodedBase64}\" {hashAttribute}=\"{encodedHash}\" {configAttribute}=\"{encodedBase64}\"></{elementName}>";
    }

    private static string? TryBuildDataViewHtml(string? rawContent) {
        if (string.IsNullOrWhiteSpace(rawContent)) {
            return null;
        }

        try {
            var raw = rawContent!;
            using var doc = JsonDocument.Parse(raw);
            var root = doc.RootElement;
            if (root.ValueKind != JsonValueKind.Object) {
                return null;
            }

            if (!root.TryGetProperty("rows", out var rowsElement) || rowsElement.ValueKind != JsonValueKind.Array) {
                return null;
            }

            var rows = new List<IReadOnlyList<string>>();
            foreach (var rowElement in rowsElement.EnumerateArray()) {
                if (rowElement.ValueKind != JsonValueKind.Array) {
                    return null;
                }

                var row = new List<string>();
                foreach (var cellElement in rowElement.EnumerateArray()) {
                    row.Add(ReadJsonScalar(cellElement));
                }

                rows.Add(row);
            }

            if (rows.Count == 0) {
                return null;
            }

            var kind = root.TryGetProperty("kind", out var kindElement) ? kindElement.GetString() : null;
            var callId = root.TryGetProperty("call_id", out var callIdElement) ? callIdElement.GetString() : null;

            return BuildDataViewHtml(rows, kind, callId);
        } catch (JsonException) {
            return null;
        }
    }

    private static string BuildDataViewHtml(IReadOnlyList<IReadOnlyList<string>> rows, string? kind, string? callId) {
        var sb = new StringBuilder();
        sb.Append("<div class=\"omd-dataview\"");
        if (!string.IsNullOrWhiteSpace(kind)) {
            sb.Append(" data-ix-kind=\"")
              .Append(System.Net.WebUtility.HtmlEncode(kind))
              .Append('"');
        }
        if (!string.IsNullOrWhiteSpace(callId)) {
            sb.Append(" data-ix-call-id=\"")
              .Append(System.Net.WebUtility.HtmlEncode(callId))
              .Append('"');
        }
        sb.Append("><table class=\"omd-dataview-table\">");

        var headers = rows[0];
        sb.Append("<thead><tr>");
        for (int i = 0; i < headers.Count; i++) {
            sb.Append("<th>")
              .Append(System.Net.WebUtility.HtmlEncode(headers[i] ?? string.Empty))
              .Append("</th>");
        }
        sb.Append("</tr></thead>");

        if (rows.Count > 1) {
            sb.Append("<tbody>");
            for (int rowIndex = 1; rowIndex < rows.Count; rowIndex++) {
                var row = rows[rowIndex];
                sb.Append("<tr>");
                for (int cellIndex = 0; cellIndex < headers.Count; cellIndex++) {
                    var cellValue = cellIndex < row.Count ? row[cellIndex] : string.Empty;
                    sb.Append("<td>")
                      .Append(System.Net.WebUtility.HtmlEncode(cellValue ?? string.Empty))
                      .Append("</td>");
                }
                sb.Append("</tr>");
            }
            sb.Append("</tbody>");
        }

        sb.Append("</table></div>");
        return sb.ToString();
    }

    private static string ReadJsonScalar(JsonElement element) {
        return element.ValueKind switch {
            JsonValueKind.String => element.GetString() ?? string.Empty,
            JsonValueKind.Number => element.GetRawText(),
            JsonValueKind.True => "true",
            JsonValueKind.False => "false",
            JsonValueKind.Null => string.Empty,
            _ => element.GetRawText()
        };
    }

    private static string? BuildNetworkShellHeadHtml(MarkdownRendererOptions options, AssetMode assetMode) {
        var network = options.Network;
        if (network?.Enabled != true) {
            return null;
        }

        var sb = new StringBuilder(512);

        var cssUrl = ResolveCssHref(network.CssUrl, assetMode);
        if (!string.IsNullOrWhiteSpace(cssUrl)) {
            sb.Append("\n<link rel=\"stylesheet\" href=\"")
              .Append(System.Net.WebUtility.HtmlEncode(cssUrl))
              .Append("\">\n");
        }

        sb.Append("""
<style>
.omd-network {
  min-height: 320px;
  margin: 1rem 0;
}
.omd-network-canvas {
  width: 100%;
  min-height: 320px;
  border: 1px solid rgba(127, 127, 127, .2);
  border-radius: 12px;
  background: rgba(127, 127, 127, .04);
}
</style>
""");

        var scriptUrl = ResolveScriptSrc(network.ScriptUrl, assetMode);
        if (!string.IsNullOrWhiteSpace(scriptUrl)) {
            sb.Append("\n<script defer src=\"")
              .Append(System.Net.WebUtility.HtmlEncode(scriptUrl))
              .Append("\"></script>\n");
        }

        return sb.ToString();
    }

    private static string? BuildNetworkBeforeReplaceScript(MarkdownRendererOptions options) {
        if (options.Network?.Enabled != true) {
            return null;
        }

        return """
try {
  root.querySelectorAll('.omd-network').forEach(host => {
    try {
      if (host.__omdVisNetwork && typeof host.__omdVisNetwork.destroy === 'function') {
        host.__omdVisNetwork.destroy();
      }
    } catch(e) { /* ignore */ }

    try { delete host.__omdVisNetwork; } catch(_) { host.__omdVisNetwork = null; }
    host.removeAttribute('data-omd-visual-rendered');
    host.removeAttribute('data-network-rendered');
  });
} catch(e) { /* ignore */ }
""";
    }

    private static string? BuildNetworkAfterReplaceScript(MarkdownRendererOptions options) {
        if (options.Network?.Enabled != true) {
            return null;
        }

        return """
try {
  function omdNetworkB64ToUtf8(b64) {
    try {
      const bytes = Uint8Array.from(atob(b64), c => c.charCodeAt(0));
      if (window.TextDecoder) return new TextDecoder('utf-8').decode(bytes);
      return decodeURIComponent(escape(String.fromCharCode.apply(null, Array.from(bytes))));
    } catch(e) { return ''; }
  }

  if (window.vis && window.vis.Network) {
    root.querySelectorAll('.omd-network:not([data-network-rendered])').forEach(host => {
      const b64 = host.getAttribute('data-network-config-b64');
      if (!b64) return;

      const jsonText = omdNetworkB64ToUtf8(b64);
      if (!jsonText) return;

      let cfg = null;
      try { cfg = JSON.parse(jsonText); } catch(e) { console.warn('vis-network config JSON parse error:', e); return; }

      try {
        const nodes = Array.isArray(cfg && cfg.nodes) ? cfg.nodes : [];
        const edges = Array.isArray(cfg && cfg.edges) ? cfg.edges : [];
        const netOptions = cfg && cfg.options && typeof cfg.options === 'object' ? cfg.options : {};

        let canvas = host.querySelector('.omd-network-canvas');
        if (!canvas) {
          canvas = document.createElement('div');
          canvas.className = 'omd-network-canvas';
          host.appendChild(canvas);
        }

        const network = new window.vis.Network(canvas, { nodes: nodes, edges: edges }, netOptions);
        host.__omdVisNetwork = network;
        host.setAttribute('data-omd-visual-rendered', 'true');
        host.setAttribute('data-network-rendered', 'true');

        setTimeout(() => {
          try { network.fit({ animation: false }); } catch(_) { /* ignore */ }
        }, 0);
      } catch(e) { console.warn('vis-network render error:', e); }
    });
  }
} catch(e) { /* ignore */ }
""";
    }

    private static string ResolveScriptSrc(string? url, AssetMode assetMode) {
        var value = (url ?? string.Empty).Trim();
        if (value.Length == 0) {
            return string.Empty;
        }

        if (assetMode == AssetMode.Offline) {
            var bundled = MarkdownRenderer.BuildBundledScriptSrc(value, "application/javascript");
            if (!string.IsNullOrWhiteSpace(bundled)) {
                return bundled;
            }
        }

        return value;
    }

    private static string ResolveCssHref(string? url, AssetMode assetMode) {
        var value = (url ?? string.Empty).Trim();
        if (value.Length == 0) {
            return string.Empty;
        }

        if (assetMode == AssetMode.Offline) {
            var bundled = MarkdownRenderer.BuildBundledCssHref(value);
            if (!string.IsNullOrWhiteSpace(bundled)) {
                return bundled;
            }
        }

        return value;
    }
}
