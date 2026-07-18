using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlRenderFontFaceLoader {
    private const string ComponentName = "OfficeIMO.Html.Renderer";

    internal static OfficeFontFaceCollection Load(
        IHtmlDocument document,
        HtmlRenderResourceSet resources,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics) {
        var fonts = new OfficeFontFaceCollection();
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
        HtmlUrlPolicy resourcePolicy = HtmlResourceUrlPolicy.Create(options.GetResourceUrlPolicy());
        var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        var pipelineOptions = new HtmlResourcePipelineOptions {
            MediaContext = options.MediaContext,
            MediaWidth = options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth,
            MediaHeight = options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D
        };

        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement)
                || !HtmlComputedStyleEngine.IsApplicableMedia(
                    styleElement.GetAttribute("media") ?? string.Empty,
                    pipelineOptions.MediaContext,
                    pipelineOptions.MediaWidth!.Value,
                    pipelineOptions.MediaHeight!.Value)) {
                continue;
            }

            foreach (HtmlCssFontFaceDefinition definition in HtmlResourcePipeline.ExtractFontFaces(styleElement.TextContent, pipelineOptions)) {
                LoadDefinition(
                    definition,
                    baseUri,
                    resourcePolicy,
                    resources,
                    options,
                    diagnostics,
                    fonts,
                    reported);
            }
        }

        return fonts;
    }

    private static void LoadDefinition(
        HtmlCssFontFaceDefinition definition,
        Uri? baseUri,
        HtmlUrlPolicy resourcePolicy,
        HtmlRenderResourceSet resources,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        OfficeFontFaceCollection fonts,
        HashSet<string> reported) {
        if (definition.FamilyName.Length == 0) {
            ReportOnce(diagnostics, reported, HtmlRenderDiagnosticCodes.FontFaceInvalid, "An @font-face rule has no usable font-family descriptor.", definition.Source);
            return;
        }

        IReadOnlyList<string> sources = HtmlResourcePipeline.ExtractFontFaceUrls(definition.Source);
        OfficeFontStyle style = ResolveStyle(definition);
        foreach (string source in sources) {
            string resolved = HtmlUrlPolicyEvaluator.ResolveUrl(
                source,
                baseUri,
                resourcePolicy);
            if (resolved.Length == 0) {
                ReportOnce(diagnostics, reported, "FontResourceRejectedByPolicy", "A font face source was rejected by the configured URL policy.", source);
                continue;
            }

            byte[]? bytes = null;
            string contentType = string.Empty;
            if (resources.TryGet(source, resolved, out HtmlResolvedResource cached)) {
                bytes = cached.Bytes;
                contentType = cached.ContentType;
            } else if (resolved.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
                if (!HtmlDataUri.TryParse(resolved, out HtmlDataUri dataUri)) {
                    ReportOnce(diagnostics, reported, HtmlRenderDiagnosticCodes.FontDataUriInvalid, "A font data URI could not be decoded.", source);
                    continue;
                }

                long estimatedBytes;
                try {
                    estimatedBytes = dataUri.EstimateDecodedByteCount();
                } catch (FormatException) {
                    ReportOnce(diagnostics, reported, HtmlRenderDiagnosticCodes.FontDataUriInvalid, "A font data URI could not be decoded.", source);
                    continue;
                }

                contentType = dataUri.MediaType;

                if (!resources.CanAcceptInlineResource(estimatedBytes, options, out string diagnosticCode, out string diagnosticDetail)) {
                    ReportOnce(diagnostics, reported, diagnosticCode, "A font data URI exceeded the configured operation-wide resource budget.", source, diagnosticDetail);
                    continue;
                }

                if (!dataUri.TryDecodeBytes(out bytes)) {
                    ReportOnce(diagnostics, reported, HtmlRenderDiagnosticCodes.FontDataUriInvalid, "A font data URI could not be decoded.", source);
                    continue;
                }

                var inlineResource = new HtmlResolvedResource(bytes, contentType);
                resources.AddInline(resolved, inlineResource);
            }

            if (bytes == null) {
                continue;
            }

            if (!IsFontContentType(contentType)) {
                ReportOnce(diagnostics, reported, HtmlRenderDiagnosticCodes.ResourceContentTypeRejected, "A font face source declared an incompatible media type.", source, contentType);
                continue;
            }

            if (fonts.TryAdd(definition.FamilyName, bytes, style)) {
                return;
            }

            ReportOnce(
                diagnostics,
                reported,
                HtmlRenderDiagnosticCodes.FontFormatUnsupported,
                "A font face is not a supported TrueType glyf-outline font.",
                source,
                contentType);
        }

        ReportOnce(
            diagnostics,
            reported,
            HtmlRenderDiagnosticCodes.FontFaceUnavailable,
            "No usable source from an @font-face rule was available to the renderer.",
            definition.FamilyName,
            definition.Source);
    }

    private static OfficeFontStyle ResolveStyle(HtmlCssFontFaceDefinition definition) {
        OfficeFontStyle style = OfficeFontStyle.Regular;
        string weight = definition.Weight.Trim();
        if (string.Equals(weight, "bold", StringComparison.OrdinalIgnoreCase)
            || string.Equals(weight, "bolder", StringComparison.OrdinalIgnoreCase)
            || int.TryParse(weight, out int numericWeight) && numericWeight >= 600) {
            style |= OfficeFontStyle.Bold;
        }

        string fontStyle = definition.Style.Trim();
        if (fontStyle.StartsWith("italic", StringComparison.OrdinalIgnoreCase)
            || fontStyle.StartsWith("oblique", StringComparison.OrdinalIgnoreCase)) {
            style |= OfficeFontStyle.Italic;
        }

        return style;
    }

    private static bool IsFontContentType(string contentType) {
        string normalized = (contentType ?? string.Empty).Split(';')[0].Trim();
        return normalized.StartsWith("font/", StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith("application/font-", StringComparison.OrdinalIgnoreCase)
            || normalized.StartsWith("application/x-font-", StringComparison.OrdinalIgnoreCase)
            || string.Equals(normalized, "application/octet-stream", StringComparison.OrdinalIgnoreCase);
    }

    private static bool IsCssStyleElement(IElement styleElement) {
        string type = (styleElement.GetAttribute("type") ?? string.Empty).Split(';')[0].Trim();
        return type.Length == 0 || string.Equals(type, "text/css", StringComparison.OrdinalIgnoreCase);
    }

    private static void ReportOnce(
        HtmlDiagnosticReport diagnostics,
        HashSet<string> reported,
        string code,
        string message,
        string? source,
        string? detail = null) {
        source = NormalizeDiagnosticValue(source);
        detail = NormalizeDiagnosticValue(detail);
        string key = code + "|" + (source ?? string.Empty) + "|" + (detail ?? string.Empty);
        if (reported.Add(key)) {
            diagnostics.Add(ComponentName, code, message, HtmlDiagnosticSeverity.Warning, source, detail);
        }
    }

    private static string? NormalizeDiagnosticValue(string? value) {
        if (string.IsNullOrEmpty(value)) {
            return value;
        }

        if (value!.StartsWith("data:", StringComparison.OrdinalIgnoreCase)) {
            int comma = value.IndexOf(',');
            string prefix = comma > 0 ? value.Substring(0, Math.Min(comma, 160)) : "data:";
            return prefix + ",... (" + value.Length + " chars)";
        }

        const int maximumLength = 512;
        return value.Length <= maximumLength ? value : value.Substring(0, maximumLength) + "...";
    }
}
