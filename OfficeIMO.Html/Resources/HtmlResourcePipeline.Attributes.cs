using AngleSharp.Dom;

namespace OfficeIMO.Html;

public static partial class HtmlResourcePipeline {
    private static bool IsExecutableScriptElement(IElement element) {
        string type = (element.GetAttribute("type") ?? string.Empty).Trim();
        if (type.Length == 0) {
            return true;
        }

        int parameterStart = type.IndexOf(';');
        if (parameterStart >= 0) {
            type = type.Substring(0, parameterStart).Trim();
        }

        return string.Equals(type, "module", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "text/javascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "application/javascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "application/ecmascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "text/ecmascript", StringComparison.OrdinalIgnoreCase)
            || string.Equals(type, "text/jscript", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsHexDigit(char value) {
        return (value >= '0' && value <= '9')
            || (value >= 'a' && value <= 'f')
            || (value >= 'A' && value <= 'F');
    }

    private static bool IsEscaped(string text, int index) {
        int slashCount = 0;
        for (int i = index - 1; i >= 0 && text[i] == '\\'; i--) {
            slashCount++;
        }

        return slashCount % 2 == 1;
    }

    private static void AddSrcSet(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, Uri? baseUri, HtmlResourcePipelineOptions options) {
        string? raw = element.GetAttribute(attributeName);
        if (string.IsNullOrWhiteSpace(raw)) {
            return;
        }

        foreach (HtmlSrcSetCandidate candidate in HtmlSrcSetParser.Parse(raw, options.MaxResponsiveImageCandidates)) {
            AddRaw(manifest, kind, element, attributeName, candidate.Url, baseUri, options);
        }
    }

    private static void AddAttribute(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, Uri? baseUri, HtmlResourcePipelineOptions options, bool skipFragmentOnly = false) {
        string? source = element.GetAttribute(attributeName);
        if (skipFragmentOnly && IsFragmentOnlyReference(source)) {
            return;
        }

        if (!string.IsNullOrWhiteSpace(source)) {
            AddRaw(manifest, kind, element, attributeName, source!, baseUri, options);
        }
    }

    private static bool IsFragmentOnlyReference(string? source) {
        return !string.IsNullOrWhiteSpace(source) && source!.TrimStart().StartsWith("#", StringComparison.Ordinal);
    }

    private static void AddMetaRefresh(HtmlResourceManifest manifest, IElement element, Uri? baseUri, HtmlResourcePipelineOptions options) {
        if (!string.Equals(element.GetAttribute("http-equiv"), "refresh", StringComparison.OrdinalIgnoreCase)) {
            return;
        }

        string? content = element.GetAttribute("content");
        if (string.IsNullOrWhiteSpace(content)) {
            return;
        }

        if (!TryReadMetaRefreshUrl(content!, out string source)) {
            return;
        }

        AddRaw(manifest, HtmlResourceKind.Hyperlink, element, "content", source, baseUri, options);
    }

    private static bool TryReadMetaRefreshUrl(string content, out string source) {
        source = string.Empty;
        foreach (string parameter in SplitMetaRefreshParameters(content).Skip(1)) {
            int separator = parameter.IndexOf('=');
            if (separator <= 0) {
                continue;
            }

            string name = parameter.Substring(0, separator).Trim();
            if (!string.Equals(name, "url", StringComparison.OrdinalIgnoreCase)) {
                continue;
            }

            source = parameter.Substring(separator + 1).Trim();
            break;
        }

        if (source.Length == 0) {
            return false;
        }

        if (source.Length > 1 && ((source[0] == '"' && source[source.Length - 1] == '"') || (source[0] == '\'' && source[source.Length - 1] == '\''))) {
            source = source.Substring(1, source.Length - 2).Trim();
        }

        return source.Length > 0;
    }

    private static IEnumerable<string> SplitMetaRefreshParameters(string content) {
        int start = 0;
        char quote = '\0';
        for (int i = 0; i < content.Length; i++) {
            char current = content[i];
            if (quote != '\0') {
                if (current == quote && !IsEscaped(content, i)) {
                    quote = '\0';
                }

                continue;
            }

            if (current == '"' || current == '\'') {
                quote = current;
                continue;
            }

            if (current == ';') {
                yield return content.Substring(start, i - start).Trim();
                start = i + 1;
            }
        }

        yield return content.Substring(start).Trim();
    }

    private static void AddRaw(HtmlResourceManifest manifest, HtmlResourceKind kind, IElement element, string attributeName, string source, Uri? baseUri, HtmlResourcePipelineOptions options) {
        HtmlUrlPolicy? policy = kind == HtmlResourceKind.Hyperlink
            ? options.UrlPolicy
            : GetResourceUrlPolicy(options);
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, policy);
        bool isAllowed = !string.IsNullOrWhiteSpace(resolved) && IsResourceKindSchemeAllowed(kind, resolved);
        manifest.Add(new HtmlResourceReference(
            kind,
            element.TagName.ToLowerInvariant(),
            attributeName,
            source.Trim(),
            resolved,
            isAllowed,
            isAllowed ? string.Empty : GetDiagnosticCode(kind)));
    }

    private static string GetDiagnosticCode(HtmlResourceKind kind) {
        switch (kind) {
            case HtmlResourceKind.Image:
                return "ImageResourceRejectedByPolicy";
            case HtmlResourceKind.Stylesheet:
                return "StylesheetResourceRejectedByPolicy";
            case HtmlResourceKind.Hyperlink:
                return "HyperlinkRejectedByPolicy";
            case HtmlResourceKind.Script:
                return "ScriptResourceRejectedByPolicy";
            case HtmlResourceKind.Media:
                return "MediaResourceRejectedByPolicy";
            case HtmlResourceKind.Font:
                return "FontResourceRejectedByPolicy";
            default:
                return "HtmlResourceRejectedByPolicy";
        }
    }

    private static bool IsResourceKindSchemeAllowed(HtmlResourceKind kind, string resolved) {
        if (kind == HtmlResourceKind.Hyperlink) {
            return true;
        }

        return !Uri.TryCreate(resolved, UriKind.Absolute, out var uri)
            || !uri.Scheme.Equals(Uri.UriSchemeMailto, StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsAllowedResourceCandidate(HtmlResourceKind kind, string? source, Uri? baseUri, HtmlUrlPolicy resourcePolicy) {
        string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(source, baseUri, resourcePolicy);
        return !string.IsNullOrWhiteSpace(resolved) && IsResourceKindSchemeAllowed(kind, resolved);
    }

}
