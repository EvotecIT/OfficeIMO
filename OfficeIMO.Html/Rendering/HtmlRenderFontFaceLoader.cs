using AngleSharp.Dom;
using AngleSharp.Html.Dom;
using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

internal static class HtmlRenderFontFaceLoader {
    private const string ComponentName = "OfficeIMO.Html.Renderer";

    internal static OfficeFontFaceCollection Load(
        IHtmlDocument document,
        HtmlResourceSession resources,
        HtmlRenderOptions options,
        HtmlConversionLimits limits,
        HtmlDiagnosticReport diagnostics) {
        var fonts = new OfficeFontFaceCollection {
            FontProgramProvider = options.Fonts?.FontProgramProvider,
            FontVariationResolver = options.Fonts?.FontVariationResolver
        };
        Uri? baseUri = HtmlDocumentParser.ResolveEffectiveBaseUri(document, options.BaseUri);
        HtmlUrlPolicy resourcePolicy = HtmlResourceUrlPolicy.Create(options.GetResourceUrlPolicy());
        var reported = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        long decodedFontBytes = 0L;
        var pipelineOptions = new HtmlResourcePipelineOptions {
            Limits = limits.Clone(),
            MediaContext = options.MediaContext,
            MediaWidth = options.Mode == HtmlRenderMode.Paged ? options.PageWidth : options.ViewportWidth,
            MediaHeight = options.Mode == HtmlRenderMode.Paged ? options.PageHeight : options.ViewportHeight ?? 1056D,
            MediaFeatures = options.MediaFeatures.Clone()
        };

        foreach (IElement styleElement in document.QuerySelectorAll("style")) {
            if (!IsCssStyleElement(styleElement)
                || !HtmlComputedStyleEngine.IsApplicableMedia(
                    styleElement.GetAttribute("media") ?? string.Empty,
                    pipelineOptions.MediaContext,
                    pipelineOptions.MediaWidth!.Value,
                    pipelineOptions.MediaHeight!.Value,
                    pipelineOptions.MediaFeatures)) {
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
                    reported,
                    ref decodedFontBytes);
            }
        }

        return fonts;
    }

    private static void LoadDefinition(
        HtmlCssFontFaceDefinition definition,
        Uri? baseUri,
        HtmlUrlPolicy resourcePolicy,
        HtmlResourceSession resources,
        HtmlRenderOptions options,
        HtmlDiagnosticReport diagnostics,
        OfficeFontFaceCollection fonts,
        HashSet<string> reported,
        ref long decodedFontBytes) {
        if (definition.FamilyName.Length == 0) {
            ReportOnce(diagnostics, reported, HtmlRenderDiagnosticCodes.FontFaceInvalid, "An @font-face rule has no usable font-family descriptor.", definition.Source);
            return;
        }

        IReadOnlyList<string> sources = HtmlResourcePipeline.ExtractFontFaceUrls(definition.Source);
        OfficeFontStyle style = ResolveStyle(definition);
        OfficeFontUnicodeRangeSet ranges = OfficeFontUnicodeRangeSet.All;
        if (!string.IsNullOrWhiteSpace(definition.UnicodeRange)) {
            if (!OfficeFontUnicodeRangeSet.TryParseCss(definition.UnicodeRange, out OfficeFontUnicodeRangeSet? parsedRanges)
                || parsedRanges == null) {
                ReportOnce(
                    diagnostics,
                    reported,
                    HtmlRenderDiagnosticCodes.FontFaceInvalid,
                    "An @font-face rule has an invalid or excessive unicode-range descriptor.",
                    definition.FamilyName,
                    definition.UnicodeRange);
                return;
            }
            ranges = parsedRanges;
        }
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
                bytes = cached.EncodedBytes;
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

                if (!resources.CanAcceptInlineResource(estimatedBytes, out string diagnosticCode, out string diagnosticDetail)) {
                    ReportOnce(diagnostics, reported, diagnosticCode, "A font data URI exceeded the configured operation-wide resource budget.", source, diagnosticDetail);
                    continue;
                }

                if (!dataUri.TryDecodeBytes(out bytes)) {
                    ReportOnce(diagnostics, reported, HtmlRenderDiagnosticCodes.FontDataUriInvalid, "A font data URI could not be decoded.", source);
                    continue;
                }

                var inlineResource = new HtmlResolvedResource(bytes, contentType);
                if (!resources.TryAcceptInline(HtmlResourceKind.Font, resolved, inlineResource,
                        out diagnosticCode, out diagnosticDetail)) {
                    ReportOnce(diagnostics, reported, diagnosticCode,
                        diagnosticCode == HtmlRenderDiagnosticCodes.ResourceContentTypeRejected
                            ? "A font face source declared an incompatible media type."
                            : "A font data URI exceeded the configured operation-wide resource budget.",
                        source, diagnosticDetail);
                    bytes = null;
                    continue;
                }
            }

            if (bytes == null) {
                continue;
            }

            if (!IsFontContentType(contentType)) {
                ReportOnce(diagnostics, reported, HtmlRenderDiagnosticCodes.ResourceContentTypeRejected, "A font face source declared an incompatible media type.", source, contentType);
                continue;
            }

            long remainingDecodedBytes = resources.MaxTotalResourceBytes
                - resources.AcceptedResourceBytes
                - decodedFontBytes;
            if (remainingDecodedBytes <= 0L) {
                ReportOnce(
                    diagnostics,
                    reported,
                    HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded,
                    "Decoded font data exceeded the configured operation-wide resource budget.",
                    source,
                    "decodedFontBytes=" + decodedFontBytes);
                continue;
            }

            int maximumDecodedBytes = (int)Math.Min(remainingDecodedBytes, int.MaxValue);
            if (fonts.TryAddBounded(
                definition.FamilyName,
                bytes,
                style,
                ranges,
                maximumDecodedBytes,
                out int acceptedDecodedBytes,
                out string? fontError)) {
                decodedFontBytes += acceptedDecodedBytes;
                return;
            }

            bool decodedLimitExceeded = fontError?.IndexOf("limit", StringComparison.OrdinalIgnoreCase) >= 0;
            ReportOnce(
                diagnostics,
                reported,
                decodedLimitExceeded
                    ? HtmlRenderDiagnosticCodes.TotalResourceByteLimitExceeded
                    : HtmlRenderDiagnosticCodes.FontFormatUnsupported,
                decodedLimitExceeded
                    ? "Decoded font data exceeded the configured operation-wide resource budget."
                    : "A font face is not supported by the first-party font engine or the configured font-program provider.",
                source,
                decodedLimitExceeded
                    ? "limit=" + maximumDecodedBytes
                    : contentType);
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
        return HtmlResourcePipeline.IsCssStyleElement(styleElement);
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
