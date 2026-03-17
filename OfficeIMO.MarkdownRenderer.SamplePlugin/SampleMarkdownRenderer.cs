using System.Text.Json;
using System.Text.RegularExpressions;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Html;

namespace OfficeIMO.MarkdownRenderer.SamplePlugin;

/// <summary>
/// Sample third-party-style renderer plugin demonstrating shared visual host output
/// plus HTML round-trip hints for custom semantic fenced blocks.
/// </summary>
public static class SampleMarkdownRenderer {
    private const string StatusPanelVisualKind = "status-panel";
    private const string StatusBadgeCssClass = "sample-status-badge";
    private const string StatusPanelCaptionAttribute = "data-sample-panel-caption";
    private const string StatusPanelSummaryAttribute = "data-sample-panel-summary";
    private const string StatusPanelStatusAttribute = "data-sample-panel-status";
    private const string StatusPanelVendorPayloadAttribute = "data-sample-status-panel-json";

    /// <summary>
    /// Reader-side transform that upgrades sample inline status badge tokens into typed inline AST.
    /// It recognizes plain-text tokens in <c>{{status:Healthy}}</c> form.
    /// </summary>
    public static IMarkdownDocumentTransform StatusBadgeReaderDocumentTransform { get; } = new StatusBadgeInlineTokenTransform();

    /// <summary>
    /// HTML-ingestion transform that upgrades recovered <c>status-panel</c> fenced code blocks
    /// into semantic fenced blocks even when the source HTML did not use the shared visual contract.
    /// </summary>
    public static IMarkdownDocumentTransform StatusPanelHtmlDocumentTransform { get; } = new StatusPanelCodeBlockTransform();

    /// <summary>
    /// Vendor-specific HTML element converter for status-panel blocks that do not use the shared visual contract.
    /// </summary>
    public static HtmlElementBlockConverter StatusPanelVendorHtmlConverter { get; } = new HtmlElementBlockConverter(
        "sample.status-panel-vendor-html",
        "Sample status-panel vendor HTML",
        TryConvertVendorStatusPanelElement);

    /// <summary>
    /// Vendor-specific inline HTML converter for sample status badges.
    /// </summary>
    public static HtmlInlineElementConverter StatusBadgeInlineConverter { get; } = new HtmlInlineElementConverter(
        "sample.status-badge-inline",
        "Sample status badge inline HTML",
        TryConvertStatusBadgeInlineElement);

    /// <summary>
    /// Shared visual-host HTML round-trip hint for recovering status-panel captions from sample host metadata.
    /// </summary>
    public static MarkdownVisualElementRoundTripHint StatusPanelRoundTripHint { get; } = new MarkdownVisualElementRoundTripHint(
        "sample.status-panel-caption",
        "Sample status-panel caption",
        context => {
            if (!string.Equals(context.VisualElement.VisualKind, StatusPanelVisualKind, StringComparison.OrdinalIgnoreCase)) {
                return null;
            }

            return context.VisualElement.TryGetAttribute(StatusPanelCaptionAttribute, out var caption)
                && !string.IsNullOrWhiteSpace(caption)
                ? context.CreateBlock(caption: caption)
                : context.CreateBlock();
        });

    /// <summary>
    /// Sample status-panel renderer plugin.
    /// </summary>
    public static MarkdownRendererPlugin StatusPanelPlugin { get; } = new MarkdownRendererPlugin(
        "Sample Status Panels",
        new Func<MarkdownFencedCodeBlockRenderer>[] {
            CreateStatusPanelRenderer
        },
        readerDocumentTransforms: new[] { StatusBadgeReaderDocumentTransform },
        htmlDocumentTransforms: new[] { StatusPanelHtmlDocumentTransform },
        htmlElementBlockConverters: new[] { StatusPanelVendorHtmlConverter },
        htmlInlineElementConverters: new[] { StatusBadgeInlineConverter },
        visualElementRoundTripHints: new[] { StatusPanelRoundTripHint });

    /// <summary>
    /// Sample feature pack composed from the status-panel plugin.
    /// </summary>
    public static MarkdownRendererFeaturePack StatusPanelFeaturePack { get; } = new MarkdownRendererFeaturePack(
        "sample.status-panel-pack",
        "Sample Status Panels",
        new[] { StatusPanelPlugin });

    /// <summary>
    /// Applies the sample status-panel renderer plugin.
    /// </summary>
    public static void ApplyStatusPanels(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.ApplyPlugin(StatusPanelPlugin);
    }

    /// <summary>
    /// Returns <see langword="true"/> when the sample status-panel plugin is already applied.
    /// </summary>
    public static bool HasStatusPanels(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        return options.HasPlugin(StatusPanelPlugin);
    }

    /// <summary>
    /// Applies the sample status-panel feature pack.
    /// </summary>
    public static void ApplyStatusPanelFeaturePack(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.ApplyFeaturePack(StatusPanelFeaturePack);
    }

    /// <summary>
    /// Returns <see langword="true"/> when the sample status-panel feature pack is already applied.
    /// </summary>
    public static bool HasStatusPanelFeaturePack(MarkdownRendererOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        return options.HasFeaturePack(StatusPanelFeaturePack);
    }

    /// <summary>
    /// Registers the sample plugin's shared visual-host HTML round-trip hints on HTML-to-markdown options.
    /// </summary>
    public static void ApplyHtmlRoundTripHints(HtmlToMarkdownOptions options) {
        if (options == null) {
            throw new ArgumentNullException(nameof(options));
        }

        options.ApplyFeaturePack(StatusPanelFeaturePack);
    }

    private static MarkdownFencedCodeBlockRenderer CreateStatusPanelRenderer() {
        return new MarkdownFencedCodeBlockRenderer(
            "Sample status panel",
            new[] { "status-panel" },
            (match, _) => TryBuildStatusPanelHtml(match.RawContent, match.Language, match.FenceInfo)) {
            SemanticKind = StatusPanelVisualKind
        };
    }

    private static string? TryBuildStatusPanelHtml(string? rawContent, string language, MarkdownCodeFenceInfo? fenceInfo) {
        if (string.IsNullOrWhiteSpace(rawContent)) {
            return null;
        }

        try {
            using var document = JsonDocument.Parse(rawContent!);
            var root = document.RootElement;
            if (root.ValueKind != JsonValueKind.Object) {
                return null;
            }

            var payload = MarkdownVisualContract.CreatePayload(rawContent);
            var title = TryReadJsonString(root, "title", "name") ?? fenceInfo?.Title;
            var summary = TryReadJsonString(root, "summary", "description");
            var status = TryReadJsonString(root, "status", "state");
            var caption = TryReadJsonString(root, "caption", "footer");

            var sb = new StringBuilder();
            sb.Append("<section class=\"")
              .Append(System.Net.WebUtility.HtmlEncode(MarkdownVisualContract.ComposeCssClass("omd-visual omd-status-panel", fenceInfo)))
              .Append('"');
            MarkdownVisualContract.AppendFenceAttributes(sb, fenceInfo);
            MarkdownVisualContract.AppendCommonAttributes(sb, StatusPanelVisualKind, language, payload);
            if (!string.IsNullOrWhiteSpace(title)) {
                MarkdownVisualContract.AppendAttribute(sb, MarkdownVisualElementContract.AttributeVisualTitle, title);
            }
            if (!string.IsNullOrWhiteSpace(summary)) {
                MarkdownVisualContract.AppendAttribute(sb, StatusPanelSummaryAttribute, summary);
            }
            if (!string.IsNullOrWhiteSpace(status)) {
                MarkdownVisualContract.AppendAttribute(sb, StatusPanelStatusAttribute, status);
            }
            if (!string.IsNullOrWhiteSpace(caption)) {
                MarkdownVisualContract.AppendAttribute(sb, StatusPanelCaptionAttribute, caption);
            }
            sb.Append('>');

            if (!string.IsNullOrWhiteSpace(title)) {
                sb.Append("<header class=\"omd-status-panel-title\">")
                  .Append(System.Net.WebUtility.HtmlEncode(title))
                  .Append("</header>");
            }

            if (!string.IsNullOrWhiteSpace(summary)) {
                sb.Append("<p class=\"omd-status-panel-summary\">")
                  .Append(System.Net.WebUtility.HtmlEncode(summary))
                  .Append("</p>");
            }

            if (!string.IsNullOrWhiteSpace(status)) {
                sb.Append("<p class=\"omd-status-panel-status\">")
                  .Append(System.Net.WebUtility.HtmlEncode(status))
                  .Append("</p>");
            }

            sb.Append("</section>");
            return sb.ToString();
        } catch (JsonException) {
            return null;
        }
    }

    private static string? TryReadJsonString(JsonElement root, params string[] propertyNames) {
        for (int i = 0; i < propertyNames.Length; i++) {
            var propertyName = propertyNames[i];
            if (!root.TryGetProperty(propertyName, out var property) || property.ValueKind != JsonValueKind.String) {
                continue;
            }

            var value = property.GetString();
            if (!string.IsNullOrWhiteSpace(value)) {
                return value;
            }
        }

        return null;
    }

    private static IReadOnlyList<IMarkdownBlock>? TryConvertVendorStatusPanelElement(HtmlElementBlockConversionContext context) {
        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        var element = context.Element;
        if (!element.TagName.Equals("SECTION", StringComparison.OrdinalIgnoreCase)
            || !element.ClassList.Contains("sample-status-panel")) {
            return null;
        }

        var payload = element.GetAttribute(StatusPanelVendorPayloadAttribute);
        if (string.IsNullOrWhiteSpace(payload)) {
            return null;
        }
        var payloadText = payload!;

        string? title = null;
        string? caption = null;

        foreach (var child in element.Children) {
            if (child.TagName.Equals("HEADER", StringComparison.OrdinalIgnoreCase)) {
                title = context.NormalizeBlockText(child.TextContent);
            } else if (child.TagName.Equals("FOOTER", StringComparison.OrdinalIgnoreCase)) {
                caption = context.NormalizeBlockText(child.TextContent);
            }
        }

        title ??= element.GetAttribute("data-title");
        caption ??= element.GetAttribute("data-caption");

        try {
            using var document = JsonDocument.Parse(payloadText);
            title ??= TryReadJsonString(document.RootElement, "title", "name");
        } catch (JsonException) {
            // Keep opaque vendor payloads round-trippable even when they are not valid JSON.
        }

        var normalizedTitle = string.IsNullOrWhiteSpace(title) ? null : title;
        var infoString = normalizedTitle == null
            ? "status-panel"
            : "status-panel title=" + QuoteFenceAttributeValue(normalizedTitle);

        return new IMarkdownBlock[] {
            new SemanticFencedBlock(StatusPanelVisualKind, infoString, payloadText, string.IsNullOrWhiteSpace(caption) ? null : caption)
        };
    }

    private static string QuoteFenceAttributeValue(string value) {
        var normalized = value ?? string.Empty;
        var escaped = normalized
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"");
        return "\"" + escaped + "\"";
    }

    private static IReadOnlyList<IMarkdownInline>? TryConvertStatusBadgeInlineElement(HtmlInlineElementConversionContext context) {
        if (context == null) {
            throw new ArgumentNullException(nameof(context));
        }

        var element = context.Element;
        if (!element.TagName.Equals("SPAN", StringComparison.OrdinalIgnoreCase)
            || !element.ClassList.Contains(StatusBadgeCssClass)) {
            return null;
        }

        var childSequence = context.ConvertChildNodesToInlineSequence();
        if (childSequence.Nodes.Count == 0) {
            var normalized = context.NormalizeInlineText(element.TextContent);
            if (normalized.Length > 0) {
                childSequence.Text(normalized);
            }
        }

        if (childSequence.Nodes.Count == 0) {
            return Array.Empty<IMarkdownInline>();
        }

        return new IMarkdownInline[] {
            new HighlightSequenceInline(childSequence)
        };
    }

    private sealed class StatusBadgeInlineTokenTransform : IMarkdownDocumentTransform {
        private static readonly Regex StatusTokenRegex = new(@"\{\{status:(?<label>[^}\r\n]+)\}\}", RegexOptions.Compiled | RegexOptions.IgnoreCase);

        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (context == null) {
                throw new ArgumentNullException(nameof(context));
            }

            var rewritten = MarkdownDoc.Create();
            for (int i = 0; i < document.Blocks.Count; i++) {
                rewritten.Add(RewriteBlock(document.Blocks[i]));
            }

            return rewritten;
        }

        private static IMarkdownBlock RewriteBlock(IMarkdownBlock block) {
            return block switch {
                ParagraphBlock paragraph => RewriteParagraph(paragraph),
                HeadingBlock heading => RewriteHeading(heading),
                SummaryBlock summary => RewriteSummary(summary),
                _ => block
            };
        }

        private static ParagraphBlock RewriteParagraph(ParagraphBlock paragraph) {
            var rewritten = RewriteSequence(paragraph.Inlines, out var changed);
            return changed ? new ParagraphBlock(rewritten) : paragraph;
        }

        private static HeadingBlock RewriteHeading(HeadingBlock heading) {
            var rewritten = RewriteSequence(heading.Inlines, out var changed);
            return changed ? new HeadingBlock(heading.Level, rewritten) : heading;
        }

        private static SummaryBlock RewriteSummary(SummaryBlock summary) {
            var rewritten = RewriteSequence(summary.Inlines, out var changed);
            return changed ? new SummaryBlock(rewritten) : summary;
        }

        private static InlineSequence RewriteSequence(InlineSequence source, out bool changed) {
            changed = false;
            if (source == null || source.Nodes.Count == 0) {
                return source ?? new InlineSequence();
            }

            for (int i = 0; i < source.Nodes.Count; i++) {
                if (source.Nodes[i] is not TextRun) {
                    return source;
                }
            }

            var rewritten = new InlineSequence();
            for (int i = 0; i < source.Nodes.Count; i++) {
                AppendRewrittenText(rewritten, ((TextRun)source.Nodes[i]).Text, ref changed);
            }

            return changed ? rewritten : source;
        }

        private static void AppendRewrittenText(InlineSequence target, string? text, ref bool changed) {
            var value = text ?? string.Empty;
            if (value.Length == 0) {
                target.Text(string.Empty);
                return;
            }

            var matches = StatusTokenRegex.Matches(value);
            if (matches.Count == 0) {
                target.Text(value);
                return;
            }

            changed = true;
            var lastIndex = 0;
            for (int i = 0; i < matches.Count; i++) {
                var match = matches[i];
                if (!match.Success) {
                    continue;
                }

                if (match.Index > lastIndex) {
                    var leading = value.Substring(lastIndex, match.Index - lastIndex);
                    target.Text(TrimEndSingleSpace(leading));
                }

                var label = match.Groups["label"].Value.Trim();
                if (label.Length > 0) {
                    target.Highlight(label);
                }

                lastIndex = match.Index + match.Length;
            }

            if (lastIndex < value.Length) {
                target.Text(TrimStartSingleSpace(value.Substring(lastIndex)));
            }
        }

        private static string TrimEndSingleSpace(string value) {
            if (string.IsNullOrEmpty(value) || !value.EndsWith(" ", StringComparison.Ordinal)) {
                return value;
            }

            return value.Substring(0, value.Length - 1);
        }

        private static string TrimStartSingleSpace(string value) {
            if (string.IsNullOrEmpty(value) || !value.StartsWith(" ", StringComparison.Ordinal)) {
                return value;
            }

            return value.Substring(1);
        }
    }

    private sealed class StatusPanelCodeBlockTransform : IMarkdownDocumentTransform {
        public MarkdownDoc Transform(MarkdownDoc document, MarkdownDocumentTransformContext context) {
            if (document == null) {
                throw new ArgumentNullException(nameof(document));
            }

            if (context == null) {
                throw new ArgumentNullException(nameof(context));
            }

            var rewritten = MarkdownDoc.Create();
            for (int i = 0; i < document.Blocks.Count; i++) {
                rewritten.Add(RewriteBlock(document.Blocks[i]));
            }

            return rewritten;
        }

        private static IMarkdownBlock RewriteBlock(IMarkdownBlock block) {
            if (block is not CodeBlock codeBlock
                || !string.Equals(codeBlock.Language, "status-panel", StringComparison.OrdinalIgnoreCase)) {
                return block;
            }

            return new SemanticFencedBlock(StatusPanelVisualKind, codeBlock.Language, codeBlock.Content, codeBlock.Caption);
        }
    }
}
